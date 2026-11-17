"""Export a SharePoint folder's contents as nested JSON using Microsoft Graph.

Reads client_id, client_secret, and tenant_id from .env. Optionally reads
SHAREPOINT_URL from .env when no URL is passed on the command line.

Examples:
    python sharepoint_items_json.py "https://tenant.sharepoint.com/sites/Team/Shared%20Documents/Folder"
    python sharepoint_items_json.py --output export.json "https://tenant.sharepoint.com/:f:/r/sites/Team/Shared%20Documents/Folder?web=1"
    python sharepoint_items_json.py --direct "https://tenant.sharepoint.com/sites/Team/Shared%20Documents/Folder"

The default is recursive and saves JSON to sharepoint_items.json in the current
working directory. Use --output to choose a different file. Folder and file nodes
include last_modified (Microsoft Graph lastModifiedDateTime).
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote, unquote, urlparse

import msal
import requests
from dotenv import load_dotenv

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]


class GraphAPIError(RuntimeError):
    def __init__(self, status_code, message):
        super().__init__(f"Microsoft Graph returned HTTP {status_code}: {message}")
        self.status_code = status_code


class GraphClient:
    def __init__(self, access_token):
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
        })

    def get_json(self, url):
        response = self.session.get(url, timeout=45)
        try:
            payload = response.json()
        except ValueError:
            payload = {}

        if not response.ok:
            error = payload.get("error", {}) if isinstance(payload, dict) else {}
            message = error.get("message") or response.text or response.reason
            raise GraphAPIError(response.status_code, message)
        return payload

    def get_all_pages(self, url):
        while url:
            page = self.get_json(url)
            yield from page.get("value", [])
            url = page.get("@odata.nextLink")


def env_value(*names):
    for name in names:
        value = os.getenv(name)
        if value and value.strip():
            return value.strip()
    return None


def get_access_token():
    client_id = env_value("client_id", "CLIENT_ID")
    client_secret = env_value("client_secret", "CLIENT_SECRET")
    tenant_id = env_value("tenant_id", "TENANT_ID")
    missing = [key for key, val in (
        ("client_id", client_id),
        ("client_secret", client_secret),
        ("tenant_id", tenant_id),
    ) if not val]
    if missing:
        raise RuntimeError("Missing required .env value(s): " + ", ".join(missing))

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    token = result.get("access_token")
    if not token:
        reason = result.get("error_description") or result.get("error") or str(result)
        raise RuntimeError(f"Could not acquire Microsoft Graph token: {reason}")
    return token


def normalize_sharepoint_path(path):
    """Strip Copy Link wrappers such as /:f:/r/ and /:w:/r/ from URL paths."""
    path = re.sub(r"^/:[^/]+:/r(?=/)", "", path or "", flags=re.IGNORECASE)
    return path or "/"


def site_path_candidates(path):
    decoded = unquote(path or "").strip("/")
    if not decoded:
        return [""]

    segments = decoded.split("/")
    min_segments = 2 if segments[0].casefold() in {"sites", "teams", "personal"} else 1
    candidates = ["/" + "/".join(segments[:i]) for i in range(len(segments), min_segments - 1, -1)]
    candidates.append("")
    return list(dict.fromkeys(candidates))


def resolve_site(graph, hostname, url_path):
    for candidate in site_path_candidates(url_path):
        if candidate:
            encoded_path = quote(candidate.lstrip("/"), safe="/")
            endpoint = f"{GRAPH_ROOT}/sites/{hostname}:/{encoded_path}"
        else:
            endpoint = f"{GRAPH_ROOT}/sites/{hostname}"
        try:
            return graph.get_json(endpoint)
        except GraphAPIError as exc:
            # Graph can return either 404 or 400 when a candidate contains a
            # library/folder segment rather than a site segment. Keep trying
            # shorter candidates only for that specific path-resolution error.
            path_resolution_miss = (
                exc.status_code == 400
                and "resource not found for the segment" in str(exc).casefold()
            )
            if exc.status_code == 404 or path_resolution_miss:
                continue
            raise
    raise RuntimeError(
        "Could not resolve a SharePoint site from this URL. Try the site's normal "
        "browser URL and confirm the app has permission to access it."
    )


def list_drives(graph, site_id):
    return list(graph.get_all_pages(f"{GRAPH_ROOT}/sites/{quote(site_id, safe=',')}/drives"))


def is_same_or_child(target_path, base_path):
    target = (target_path or "").rstrip("/").casefold()
    base = (base_path or "").rstrip("/").casefold()
    return not base or target == base or target.startswith(base + "/")


def choose_drive(graph, site, input_path):
    """Return the matching document library and its relative folder path."""
    target_path = unquote(input_path or "").rstrip("/")
    matches = []
    for drive in list_drives(graph, site["id"]):
        web_url = drive.get("webUrl")
        if not web_url:
            continue
        base_path = unquote(urlparse(web_url).path or "").rstrip("/")
        if is_same_or_child(target_path, base_path):
            matches.append((len(base_path), drive, base_path))

    if matches:
        _, drive, base_path = max(matches, key=lambda entry: entry[0])
        remainder = target_path[len(base_path):].strip("/") if base_path else target_path.strip("/")
        return drive, remainder

    site_path = unquote(urlparse(site.get("webUrl", "")).path or "").rstrip("/")
    if target_path.casefold() == site_path.casefold():
        drive = graph.get_json(f"{GRAPH_ROOT}/sites/{quote(site['id'], safe=',')}/drive")
        return drive, ""

    raise RuntimeError(
        "The site was found, but no document library matched the URL path. "
        "Pass a document library or folder URL, not a SharePoint page/list-view URL."
    )


def item_by_path(graph, drive_id, relative_path):
    drive_id = quote(drive_id, safe="")
    if relative_path:
        encoded_path = quote(relative_path.strip("/"), safe="/")
        endpoint = f"{GRAPH_ROOT}/drives/{drive_id}/root:/{encoded_path}"
    else:
        endpoint = f"{GRAPH_ROOT}/drives/{drive_id}/root"
    return graph.get_json(endpoint)


def list_children(graph, drive_id, folder_path):
    drive_id = quote(drive_id, safe="")
    if folder_path:
        encoded_path = quote(folder_path.strip("/"), safe="/")
        endpoint = f"{GRAPH_ROOT}/drives/{drive_id}/root:/{encoded_path}:/children"
    else:
        endpoint = f"{GRAPH_ROOT}/drives/{drive_id}/root/children"
    items = list(graph.get_all_pages(endpoint))
    # Stable, readable order: folders first, then files/items alphabetically.
    items.sort(key=lambda item: (0 if "folder" in item else 1, item.get("name", "").casefold()))
    return items


def make_node(item, path, recursive, graph, drive_id):
    is_folder = "folder" in item
    is_file = "file" in item
    node = {
        "name": item.get("name", "(unnamed item)"),
        "type": "folder" if is_folder else "file" if is_file else "item",
        "path": path,
        "id": item.get("id"),
        "last_modified": item.get("lastModifiedDateTime"),
        "web_url": item.get("webUrl"),
    }

    if is_file:
        node["size_bytes"] = item.get("size")
        file_info = item.get("file") or {}
        if file_info.get("mimeType"):
            node["mime_type"] = file_info["mimeType"]

    if is_folder:
        folder_info = item.get("folder") or {}
        if folder_info.get("childCount") is not None:
            node["child_count"] = folder_info["childCount"]
        node["children"] = []
        if recursive:
            child_items = list_children(graph, drive_id, path)
            for child in child_items:
                child_path = "/".join(part for part in (path, child.get("name", "")) if part)
                node["children"].append(make_node(child, child_path, recursive, graph, drive_id))

    return node


def count_nodes(root):
    total = folders = files = 0
    def visit(node):
        nonlocal total, folders, files
        for child in node.get("children", []):
            total += 1
            if child.get("type") == "folder":
                folders += 1
                visit(child)
            elif child.get("type") == "file":
                files += 1
    visit(root)
    return {"total_items": total, "folders": folders, "files": files}


def main():
    # Load .env alongside the working directory or script location.
    load_dotenv()
    load_dotenv(Path(__file__).resolve().with_name(".env"), override=False)

    parser = argparse.ArgumentParser(
        description="Export SharePoint folder contents and modification timestamps as nested JSON."
    )
    parser.add_argument("url", nargs="?", help="SharePoint folder/library URL, or set SHAREPOINT_URL in .env")
    parser.add_argument(
        "--output", "-o",
        default="sharepoint_items.json",
        help="JSON output file (default: sharepoint_items.json in the current directory)",
    )
    parser.add_argument("--direct", action="store_true", help="Include immediate children only; don't traverse subfolders")
    args = parser.parse_args()

    sharepoint_url = args.url or env_value("SHAREPOINT_URL", "sharepoint_url")
    if not sharepoint_url:
        parser.error("provide a SharePoint URL or set SHAREPOINT_URL in .env")

    parsed = urlparse(sharepoint_url)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        parser.error("the SharePoint URL must be a full http(s) URL")

    try:
        normalized_path = normalize_sharepoint_path(parsed.path)
        graph = GraphClient(get_access_token())
        site = resolve_site(graph, parsed.hostname, normalized_path)
        drive, relative_path = choose_drive(graph, site, normalized_path)

        root_item = item_by_path(graph, drive["id"], relative_path)
        root_path = relative_path.strip("/")
        root_node = make_node(root_item, root_path, not args.direct, graph, drive["id"])
        summary = count_nodes(root_node)

        result = {
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
            "retrieved_at": datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z"),
            "recursive": not args.direct,
            "summary": summary,
            "root": root_node,
        }

        output_json = json.dumps(result, indent=2, ensure_ascii=False) + "\n"
        output_path = Path(args.output).expanduser().resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(output_json, encoding="utf-8")
        print(f"Saved JSON to {output_path}")

    except (GraphAPIError, RuntimeError, requests.RequestException) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        if isinstance(exc, GraphAPIError) and exc.status_code == 403:
            print(
                "Check that Microsoft Graph application permissions (such as Sites.Read.All) "
                "have admin consent and that the app can access this site.",
                file=sys.stderr,
            )
        sys.exit(1)


if __name__ == "__main__":
    main()
