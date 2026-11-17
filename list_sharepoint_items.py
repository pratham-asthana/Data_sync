"""List files and folders from a SharePoint site/library/folder URL using Microsoft Graph.

Credentials are read from .env: client_id, client_secret, tenant_id.
Usage:
    python list_sharepoint_items.py "https://tenant.sharepoint.com/sites/Team/Shared%20Documents/Folder"
    python list_sharepoint_items.py --direct "https://tenant.sharepoint.com/sites/Team/Shared%20Documents/Folder"
    python list_sharepoint_items.py --details --show-urls "https://tenant.sharepoint.com/sites/Team/Shared%20Documents/Folder"

You may also set SHAREPOINT_URL in .env and omit the positional URL.
"""

import argparse
import json
import os
import re
import sys
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
        """Yield every item from a Microsoft Graph collection, following nextLink."""
        while url:
            page = self.get_json(url)
            yield from page.get("value", [])
            url = page.get("@odata.nextLink")


def get_env_value(*names):
    for name in names:
        value = os.getenv(name)
        if value:
            return value.strip()
    return None


def get_access_token():
    # Accept either lowercase or uppercase .env keys.
    client_id = get_env_value("client_id", "CLIENT_ID")
    client_secret = get_env_value("client_secret", "CLIENT_SECRET")
    tenant_id = get_env_value("tenant_id", "TENANT_ID")

    missing = [name for name, value in (
        ("client_id", client_id),
        ("client_secret", client_secret),
        ("tenant_id", tenant_id),
    ) if not value]
    if missing:
        raise RuntimeError(
            "Missing required .env value(s): " + ", ".join(missing)
        )

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    token = result.get("access_token")
    if not token:
        description = result.get("error_description") or result.get("error") or str(result)
        raise RuntimeError(f"Could not acquire Microsoft Graph token: {description}")
    return token


def normalize_sharepoint_path(path):
    """Remove SharePoint's /:f:/r/, /:w:/r/, etc. sharing wrapper from a URL path.

    Links such as https://tenant.sharepoint.com/:f:/r/sites/Site/Shared%20Documents/Folder
    contain the real server-relative path after the /:f:/r/ wrapper. Query parameters
    on these links are not needed for path-based Microsoft Graph lookup.
    """
    normalized = path or ""
    normalized = re.sub(r"^/:[^/]+:/r(?=/)", "", normalized, flags=re.IGNORECASE)
    return normalized or "/"


def site_path_candidates(path):
    """Create longest-to-shortest candidate site paths from a SharePoint URL path."""
    decoded = unquote(path or "").strip("/")
    if not decoded:
        return [""]

    segments = decoded.split("/")
    if segments[0].casefold() in {"sites", "teams", "personal"}:
        # A site collection typically starts at /sites/<name> or /teams/<name>.
        min_segments = 2
    else:
        # Also allow a root site with a path-based subsite.
        min_segments = 1

    candidates = [
        "/" + "/".join(segments[:i])
        for i in range(len(segments), min_segments - 1, -1)
    ]
    candidates.append("")  # Tenant root site fallback.
    # Preserve order while removing duplicates.
    return list(dict.fromkeys(candidates))


def resolve_site(graph, hostname, url_path):
    """Resolve a site, trying longer URL path prefixes before shorter ones."""
    for candidate in site_path_candidates(url_path):
        if candidate:
            encoded_path = quote(candidate.lstrip("/"), safe="/")
            endpoint = f"{GRAPH_ROOT}/sites/{hostname}:/{encoded_path}"
        else:
            endpoint = f"{GRAPH_ROOT}/sites/{hostname}"

        try:
            site = graph.get_json(endpoint)
            return site
        except GraphAPIError as exc:
            if exc.status_code == 404:
                continue
            raise

    raise RuntimeError(
        "Could not resolve a SharePoint site from this URL. Use the site's normal "
        "browser URL (not a sharing/checkout link) and confirm the app has access."
    )


def path_is_same_or_child(target_path, base_path):
    target = (target_path or "").rstrip("/")
    base = (base_path or "").rstrip("/")
    if not base:
        return True
    return target.casefold() == base.casefold() or target.casefold().startswith(base.casefold() + "/")


def list_drives(graph, site_id):
    endpoint = f"{GRAPH_ROOT}/sites/{quote(site_id, safe=',')}/drives"
    return list(graph.get_all_pages(endpoint))


def choose_drive(graph, site, input_path):
    """Find the document library whose webUrl contains the supplied URL path."""
    drives = list_drives(graph, site["id"])
    target_path = unquote(input_path or "").rstrip("/")
    matches = []

    for drive in drives:
        web_url = drive.get("webUrl")
        if not web_url:
            continue
        base_path = unquote(urlparse(web_url).path or "").rstrip("/")
        if path_is_same_or_child(target_path, base_path):
            matches.append((len(base_path), drive, base_path))

    if matches:
        # Prefer the deepest library URL if libraries have nested URL paths.
        _, drive, base_path = max(matches, key=lambda item: item[0])
        remainder = target_path[len(base_path):].strip("/") if base_path else target_path.strip("/")
        return drive, remainder

    # If the supplied URL is exactly the site URL (rather than a library/folder URL),
    # list the root of the site's default document library.
    site_path = unquote(urlparse(site.get("webUrl", "")).path or "").rstrip("/")
    if target_path.casefold() == site_path.casefold():
        endpoint = f"{GRAPH_ROOT}/sites/{quote(site['id'], safe=',')}/drive"
        return graph.get_json(endpoint), ""

    raise RuntimeError(
        "The URL resolved to a SharePoint site, but no document library matched the "
        "URL path. Make sure the URL points to a document library or folder, rather "
        "than a SharePoint page or list view."
    )


def list_children(graph, drive_id, relative_folder_path):
    drive_id_encoded = quote(drive_id, safe="")
    if relative_folder_path:
        folder_path = quote(relative_folder_path.strip("/"), safe="/")
        endpoint = f"{GRAPH_ROOT}/drives/{drive_id_encoded}/root:/{folder_path}:/children"
    else:
        endpoint = f"{GRAPH_ROOT}/drives/{drive_id_encoded}/root/children"

    return list(graph.get_all_pages(endpoint))


def collect_items(graph, drive_id, start_path, recursive=False):
    """Collect direct children, or recursively walk all descendant folders."""
    results = []

    def walk(folder_path="", display_prefix=""):
        for item in list_children(graph, drive_id, folder_path):
            name = item.get("name", "(unnamed item)")
            display_path = f"{display_prefix}/{name}".strip("/")
            item_type = "Folder" if "folder" in item else "File" if "file" in item else "Item"

            results.append({
                "path": display_path,
                "name": name,
                "type": item_type,
                "size_bytes": item.get("size"),
                "last_modified": item.get("lastModifiedDateTime"),
                "web_url": item.get("webUrl"),
                "id": item.get("id"),
            })

            if recursive and "folder" in item:
                child_relative_path = "/".join(part for part in (folder_path, name) if part)
                walk(child_relative_path, display_path)

    walk(start_path, "")
    return results


def main():
    load_dotenv()

    parser = argparse.ArgumentParser(
        description="List files and folders from a SharePoint site/library/folder URL."
    )
    parser.add_argument(
        "url",
        nargs="?",
        help="SharePoint site, document-library, or folder URL (or set SHAREPOINT_URL in .env).",
    )
    scope = parser.add_mutually_exclusive_group()
    scope.add_argument(
        "--recursive",
        dest="recursive",
        action="store_true",
        help="Include items inside all descendant folders (the default).",
    )
    scope.add_argument(
        "--direct",
        dest="recursive",
        action="store_false",
        help="List only the selected folder's immediate children.",
    )
    parser.set_defaults(recursive=True)
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print item metadata as JSON instead of the folder tree.",
    )
    parser.add_argument(
        "--details",
        action="store_true",
        help="Show file sizes and last-modified timestamps in the tree.",
    )
    parser.add_argument(
        "--show-urls",
        action="store_true",
        help="Print each item's SharePoint web URL on a secondary line.",
    )
    args = parser.parse_args()

    sharepoint_url = args.url or get_env_value("SHAREPOINT_URL", "sharepoint_url")
    if not sharepoint_url:
        parser.error('provide a SharePoint URL or set SHAREPOINT_URL in .env')

    parsed = urlparse(sharepoint_url)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        parser.error("the SharePoint URL must be a full http(s) URL")

    try:
        # SharePoint's browser Copy Link URLs may include a wrapper such as /:f:/r/.
        # Strip that wrapper before resolving the site and document library.
        normalized_path = normalize_sharepoint_path(parsed.path)
        graph = GraphClient(get_access_token())
        site = resolve_site(graph, parsed.hostname, normalized_path)
        drive, relative_path = choose_drive(graph, site, normalized_path)
        items = collect_items(graph, drive["id"], relative_path, recursive=args.recursive)

        if args.json:
            print(json.dumps(items, indent=2, ensure_ascii=False))
            return

        folder_name = unquote(relative_path.rstrip("/").split("/")[-1]) if relative_path else drive.get("name", "Document library")
        folder_name = folder_name or drive.get("name", "Document library")
        folders_count = sum(1 for item in items if item["type"] == "Folder")
        files_count = sum(1 for item in items if item["type"] == "File")

        print(f"Site:    {site.get('webUrl', site.get('displayName', ''))}")
        print(f"Library: {drive.get('name', '(unnamed library)')}")
        print(f"Folder:  {relative_path or '/ (library root)'}")
        print(
            f"Items:   {len(items)} total — {folders_count} folder(s), "
            f"{files_count} file(s)" + (" (recursive)" if args.recursive else " (direct children)")
        )
        print()

        if not items:
            print(f"📁 {folder_name}/")
            print("└── (empty folder)")
            return

        # Rebuild the flat API results into a tree so the console reflects the
        # actual folder hierarchy rather than showing slash-separated paths.
        root = {"item": None, "children": {}}
        for item in items:
            parts = item["path"].split("/")
            node = root
            for part in parts:
                node = node["children"].setdefault(part, {"item": None, "children": {}})
            node["item"] = item

        print(f"📁 {folder_name}/")

        def pretty_size(size):
            if size is None:
                return ""
            value = float(size)
            for unit in ("B", "KB", "MB", "GB", "TB"):
                if value < 1024 or unit == "TB":
                    return f"{int(value)} {unit}" if unit == "B" else f"{value:.1f} {unit}"
                value /= 1024
            return f"{value:.1f} TB"

        def sort_key(entry):
            name, node = entry
            item = node.get("item") or {}
            # Folders first, then alphabetical order.
            return (0 if item.get("type") == "Folder" else 1, name.casefold())

        def render_tree(node, prefix=""):
            entries = sorted(node["children"].items(), key=sort_key)
            for index, (name, child) in enumerate(entries):
                last = index == len(entries) - 1
                branch = "└── " if last else "├── "
                item = child.get("item") or {}
                item_type = item.get("type", "Folder" if child["children"] else "Item")
                icon = "📁" if item_type == "Folder" else "📄" if item_type == "File" else "📦"
                suffix = "/" if item_type == "Folder" else ""
                details = ""
                if args.details:
                    bits = []
                    if item_type == "File" and item.get("size_bytes") is not None:
                        bits.append(pretty_size(item["size_bytes"]))
                    if item.get("last_modified"):
                        bits.append("modified " + item["last_modified"].replace("T", " ").replace("Z", " UTC"))
                    if bits:
                        details = "  (" + "; ".join(bits) + ")"
                print(f"{prefix}{branch}{icon} {name}{suffix}{details}")

                child_prefix = prefix + ("    " if last else "│   ")
                if args.show_urls and item.get("web_url"):
                    print(f"{child_prefix}↳ {item['web_url']}")
                if child["children"]:
                    render_tree(child, child_prefix)

        render_tree(root)

    except (GraphAPIError, RuntimeError, requests.RequestException) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        if isinstance(exc, GraphAPIError) and exc.status_code == 403:
            print(
                "Check that Microsoft Graph application permissions (for example, "
                "Sites.Read.All) have admin consent and that the app is permitted "
                "to access this site.",
                file=sys.stderr,
            )
        sys.exit(1)


if __name__ == "__main__":
    main()
