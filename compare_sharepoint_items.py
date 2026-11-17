"""Compare a fresh SharePoint listing with a previous JSON snapshot.

Detects both added and deleted files/folders.

Outputs:
    - sharepoint_items.json: Complete current SharePoint snapshot.
    - sharepoint_new_items.json: Added and deleted items, preserving hierarchy.

Existing folders are included with status="unchanged" only when needed to
preserve the hierarchy of changed items.

Example:
    python compare_sharepoint_items.py "https://tenant.sharepoint.com/sites/Team/Shared%20Documents/Folder" \
        --previous sharepoint_items.json

By default, the refreshed snapshot replaces the previous snapshot after the
comparison. Use --current-output to preserve the previous file.
"""

import argparse
import json
import sys
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse

import requests
from dotenv import load_dotenv

import sharepoint_items_json_with_date as exporter


def utc_now():
    return (
        datetime.now(timezone.utc)
        .isoformat(timespec="seconds")
        .replace("+00:00", "Z")
    )


def load_json(path: Path):
    try:
        with path.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
    except FileNotFoundError as exc:
        raise RuntimeError(
            f"Previous snapshot not found: {path}\n"
            "Create the baseline first with "
            "sharepoint_items_json_with_date.py, or pass the correct "
            "--previous path."
        ) from exc
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"Previous snapshot is not valid JSON: {path}: {exc}"
        ) from exc

    if not isinstance(data, dict) or not isinstance(data.get("root"), dict):
        raise RuntimeError(
            f"{path} does not look like a SharePoint snapshot: "
            "expected an object with a 'root' object."
        )

    return data


def item_key(node):
    """Prefer stable Graph item IDs; fall back to case-insensitive paths."""
    item_id = node.get("id")

    if item_id:
        return "id:" + str(item_id)

    path = node.get("path") or node.get("name") or ""
    return "path:" + str(path).replace("\\", "/").strip("/").casefold()


def flatten_tree(root):
    """Return every descendant keyed by its item identity."""
    items = {}

    def visit(node):
        for child in node.get("children", []) or []:
            items[item_key(child)] = child

            if child.get("type") == "folder":
                visit(child)

    visit(root)
    return items


def count_status(root, status):
    """Count added or deleted items by type."""
    total = folders = files = 0

    def visit(node):
        nonlocal total, folders, files

        for child in node.get("children", []) or []:
            if child.get("status") == status:
                total += 1

                if child.get("type") == "folder":
                    folders += 1
                elif child.get("type") == "file":
                    files += 1

            visit(child)

    visit(root)

    return {
        "total_items": total,
        "folders": folders,
        "files": files,
    }


def build_changes_tree(current_root, previous_root):
    """Build a tree of added/deleted items plus required parent folders."""
    previous_items = flatten_tree(previous_root)
    current_items = flatten_tree(current_root)

    previous_keys = set(previous_items)
    current_keys = set(current_items)

    added_keys = current_keys - previous_keys
    deleted_keys = previous_keys - current_keys

    # Build the change lookup. Deleted nodes come from the previous snapshot;
    # added nodes come from the current snapshot.
    changed_items = {}

    for key in added_keys:
        changed_items[key] = (
            deepcopy(current_items[key]),
            "added",
        )

    for key in deleted_keys:
        changed_items[key] = (
            deepcopy(previous_items[key]),
            "deleted",
        )

    def node_has_changes(node, source_items, source_keys):
        """Check whether a folder contains a changed descendant."""
        for child in node.get("children", []) or []:
            child_key = item_key(child)

            if child_key in changed_items:
                return True

            if child.get("type") == "folder":
                if node_has_changes(child, source_items, source_keys):
                    return True

        return False

    def build_node(node, source, is_root=False):
        key = item_key(node)

        if is_root:
            status = "unchanged"
        elif key in changed_items:
            status = changed_items[key][1]
            node = changed_items[key][0]
        else:
            status = "unchanged"

        result = {
            k: deepcopy(v)
            for k, v in node.items()
            if k != "children"
        }

        if not is_root:
            result["status"] = status
            result["is_new"] = status == "added"

        if node.get("type") == "folder" or "children" in node:
            result["children"] = []

            # For added items, use the current tree; for deleted items,
            # use the previous tree. For unchanged parent folders, prefer
            # the current tree so additions are placed in the current layout.
            if status == "deleted":
                child_source = previous_items
                children = node.get("children", []) or []
            else:
                children = node.get("children", []) or []

            for child in children:
                child_key = item_key(child)

                if child_key in changed_items:
                    child_node = build_node(
                        changed_items[child_key][0],
                        source,
                    )
                    result["children"].append(child_node)

                elif child.get("type") == "folder":
                    # Existing folders are retained only when they contain
                    # a changed descendant.
                    child_node = build_node(child, source)

                    if child_node is not None:
                        result["children"].append(child_node)

            # Deleted folders need their previous descendants, including
            # items that disappeared together with the folder.
            if status == "deleted":
                result["children"] = []

                for child in node.get("children", []) or []:
                    child_key = item_key(child)

                    if child_key in deleted_keys:
                        result["children"].append(
                            build_node(
                                previous_items[child_key],
                                "previous",
                            )
                        )
                    elif child.get("type") == "folder":
                        child_node = build_node(child, "previous")

                        if child_node is not None:
                            result["children"].append(child_node)

            if not is_root and status == "unchanged":
                if not result["children"]:
                    return None

        elif status == "unchanged" and not is_root:
            return None

        return result

    # Construct the result starting from the current root. Added items are
    # represented at their current locations. Deleted items that no longer
    # have a current parent are attached to their last known parent where
    # possible, or to the root if that parent was deleted.
    root = build_node(current_root, "current", is_root=True)

    included_deleted = set()

    def record_deleted(node):
        if node.get("status") == "deleted":
            included_deleted.add(item_key(node))

        for child in node.get("children", []) or []:
            record_deleted(child)

    record_deleted(root)

    # If deleted items were not reachable through the current tree, place
    # them using their former parent relationships.
    deleted_by_parent = {}

    for key in deleted_keys:
        node = previous_items[key]
        parent_path = str(node.get("path", "")).replace("\\", "/").strip("/")
        parent_path = parent_path.rsplit("/", 1)[0] if "/" in parent_path else ""

        deleted_by_parent.setdefault(parent_path.casefold(), []).append(node)

    def find_node_by_path(node, path):
        if str(node.get("path", "")).replace("\\", "/").strip("/").casefold() == path:
            return node

        for child in node.get("children", []) or []:
            found = find_node_by_path(child, path)

            if found is not None:
                return found

        return None

    for key in deleted_keys:
        if key in included_deleted:
            continue

        node = deepcopy(previous_items[key])
        node["status"] = "deleted"
        node["is_new"] = False
        node.pop("children", None)

        path = str(node.get("path", "")).replace("\\", "/").strip("/")
        parent_path = path.rsplit("/", 1)[0] if "/" in path else ""

        parent = find_node_by_path(root, parent_path.casefold())

        if parent is None:
            parent = root

        parent.setdefault("children", []).append(node)

    return root, added_keys, deleted_keys


def fetch_current_snapshot(sharepoint_url, recursive=True):
    parsed = urlparse(sharepoint_url)

    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        raise RuntimeError("The SharePoint URL must be a full http(s) URL.")

    normalized_path = exporter.normalize_sharepoint_path(parsed.path)
    graph = exporter.GraphClient(exporter.get_access_token())

    site = exporter.resolve_site(
        graph,
        parsed.hostname,
        normalized_path,
    )

    drive, relative_path = exporter.choose_drive(
        graph,
        site,
        normalized_path,
    )

    root_item = exporter.item_by_path(
        graph,
        drive["id"],
        relative_path,
    )

    root_path = relative_path.strip("/")

    root_node = exporter.make_node(
        root_item,
        root_path,
        recursive,
        graph,
        drive["id"],
    )

    return {
        "site": {
            "name": site.get("displayName"),
            "url": site.get("webUrl"),
            "id": site.get("id"),
        },
        "library": {
            "name": drive.get("name"),
            "id": drive.get("id"),
            "web_url": drive.get("webUrl"),
        },
        "retrieved_at": utc_now(),
        "recursive": recursive,
        "summary": exporter.count_nodes(root_node),
        "root": root_node,
    }


def validate_same_snapshot_scope(previous, current):
    """Refuse to compare different sites, libraries, or folder roots."""
    checks = (
        (
            "site.id",
            previous.get("site", {}).get("id"),
            current.get("site", {}).get("id"),
        ),
        (
            "library.id",
            previous.get("library", {}).get("id"),
            current.get("library", {}).get("id"),
        ),
        (
            "root.id",
            previous.get("root", {}).get("id"),
            current.get("root", {}).get("id"),
        ),
    )

    for label, old_value, new_value in checks:
        if old_value and new_value and old_value != new_value:
            raise RuntimeError(
                f"The previous snapshot and current URL refer to different "
                f"{label} values. Check --previous and the SharePoint URL; "
                "the baseline was not overwritten."
            )


def write_json(path: Path, payload):
    path = path.expanduser().resolve()
    path.parent.mkdir(parents=True, exist_ok=True)

    path.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )

    return path


def main():
    # Load .env from the current directory and alongside this script.
    load_dotenv()
    load_dotenv(
        Path(__file__).resolve().with_name(".env"),
        override=False,
    )

    parser = argparse.ArgumentParser(
        description=(
            "Refresh a SharePoint snapshot and detect added/deleted items."
        )
    )

    parser.add_argument(
        "url",
        nargs="?",
        help="SharePoint folder/library URL, or set SHAREPOINT_URL in .env",
    )

    parser.add_argument(
        "--previous",
        default="sharepoint_items.json",
        help="Existing baseline JSON (default: sharepoint_items.json)",
    )

    parser.add_argument(
        "--current-output",
        default=None,
        help=(
            "Where to save the refreshed full snapshot "
            "(default: overwrite --previous after comparison)"
        ),
    )

    parser.add_argument(
        "--new-output",
        default="sharepoint_new_items.json",
        help=(
            "Where to save the JSON of changes "
            "(default: sharepoint_new_items.json)"
        ),
    )

    parser.add_argument(
        "--direct",
        action="store_true",
        help="List only immediate children; default traverses all subfolders",
    )

    args = parser.parse_args()

    sharepoint_url = (
        args.url
        or exporter.env_value("SHAREPOINT_URL", "sharepoint_url")
    )

    if not sharepoint_url:
        parser.error(
            "Provide a SharePoint URL or set SHAREPOINT_URL in .env."
        )

    previous_path = Path(args.previous).expanduser().resolve()

    current_output_path = (
        Path(args.current_output).expanduser()
        if args.current_output
        else previous_path
    )

    new_output_path = Path(args.new_output).expanduser()

    if current_output_path.resolve() == new_output_path.resolve():
        parser.error(
            "--current-output and --new-output must be different files."
        )

    try:
        # Load the baseline before writing the refreshed snapshot.
        previous = load_json(previous_path)

        current = fetch_current_snapshot(
            sharepoint_url,
            recursive=not args.direct,
        )

        validate_same_snapshot_scope(previous, current)

        changes_root, added_keys, deleted_keys = build_changes_tree(
            current["root"],
            previous["root"],
        )

        added_summary = count_status(changes_root, "added")
        deleted_summary = count_status(changes_root, "deleted")

        # Counts are derived directly from the compared item identities.
        added_items = {
            item_key(node): node
            for node in flatten_tree(current["root"]).values()
        }

        deleted_items = {
            item_key(node): node
            for node in flatten_tree(previous["root"]).values()
        }

        added_folders = sum(
            1 for key in added_keys
            if added_items[key].get("type") == "folder"
        )
        added_files = sum(
            1 for key in added_keys
            if added_items[key].get("type") == "file"
        )

        deleted_folders = sum(
            1 for key in deleted_keys
            if deleted_items[key].get("type") == "folder"
        )
        deleted_files = sum(
            1 for key in deleted_keys
            if deleted_items[key].get("type") == "file"
        )

        new_result = {
            "site": current["site"],
            "library": current["library"],
            "compared_at": utc_now(),
            "previous_snapshot": str(previous_path),
            "current_snapshot": str(
                current_output_path.expanduser().resolve()
            ),
            "summary": {
                "total_changes": len(added_keys) + len(deleted_keys),
                "added": {
                    "total_items": len(added_keys),
                    "folders": added_folders,
                    "files": added_files,
                },
                "deleted": {
                    "total_items": len(deleted_keys),
                    "folders": deleted_folders,
                    "files": deleted_files,
                },
            },
            "root": changes_root,
        }

        saved_current = write_json(current_output_path, current)
        saved_changes = write_json(new_output_path, new_result)

        if added_keys or deleted_keys:
            print(
                f"Changes detected: {len(added_keys)} added, "
                f"{len(deleted_keys)} deleted."
            )
        else:
            print("No added or deleted files or folders found.")

        print(f"Full current snapshot: {saved_current}")
        print(f"Changes JSON:          {saved_changes}")
        print(
            'In the changes JSON, status is "added", "deleted", '
            'or "unchanged". Unchanged folders are included only to '
            "preserve the hierarchy."
        )

    except (
        exporter.GraphAPIError,
        RuntimeError,
        requests.RequestException,
    ) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)

        if (
            isinstance(exc, exporter.GraphAPIError)
            and exc.status_code == 403
        ):
            print(
                "Check Microsoft Graph application permissions "
                "(for example, Sites.Read.All), admin consent, "
                "and the app's access to this site.",
                file=sys.stderr,
            )

        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())