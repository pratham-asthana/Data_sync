"""Export structures and document metadata for newly discovered SharePoint sites.

Run this script to execute comparison.py, then export each added site's folder tree
and a JSON list of documents under the requested delivery paths. Credentials are
loaded from the project's .env file.

Output JSON matches the Neo4j schema exactly for direct import:
- Properties: department, document_name, document_type, file_extension,
  modified_timestamp, sharepoint_path, sharepoint_url, version
- No summary field (excluded per requirement)
- version defaults to "UNKNOWN" when not found
- document_type values match existing Neo4j node labels
"""

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path
from urllib.parse import urlparse

from dotenv import load_dotenv

from sharepoint_items_json_with_date import (
    GraphClient,
    get_access_token,
    list_children,
    list_drives,
    resolve_site,
)

BASE_DIR = Path(__file__).resolve().parent
COMPARISON_SCRIPT = BASE_DIR / "comparison.py"
DEFAULT_OUTPUT_DIR = BASE_DIR / "new_site_exports"

# Track folder name (lowercase) -> Department value (matches Neo4j data)
TRACKS = {
    "epm": "EPM",
    "fin": "FIN",
    "hcm": "HCM",
    "scm": "SCM",
    "extensions": "Extensions",
}

# Folder path patterns -> Neo4j document_type (must match existing node labels)
# Format: (("folder1", "folder2"), "NEO4J_LABEL")
DOCUMENT_PATHS = (
    (("understudy", "rtm"), "RTM"),
    (("redesign", "config workbooks"), "CONFIG_WORKBOOK"),
    (("redesign", "sdd"), "SDD"),
    (("orchestrate", "kut"), "USER_GUIDE"),      # KUT recordings -> USER_GUIDE
    (("orchestrate", "uat"), "USER_GUIDE"),      # UAT recordings -> USER_GUIDE
)

# Version pattern: v1.0, v2.3, v1.2.3 etc (case insensitive, not part of larger word)
VERSION_PATTERN = re.compile(r"(?<![A-Za-z0-9])v\d+(?:\.\d+)+(?!\d)", re.IGNORECASE)


def run_comparison():
    result = subprocess.run(
        [sys.executable, str(COMPARISON_SCRIPT), "--json"],
        cwd=BASE_DIR,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        details = result.stderr.strip() or result.stdout.strip()
        raise RuntimeError(details or "comparison.py failed.")

    try:
        payload = json.loads(result.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            "comparison.py did not return valid JSON. Run it directly to inspect its output."
        ) from exc

    sites = payload.get("added_sites")
    if not isinstance(sites, list):
        raise RuntimeError("comparison.py result is missing the added_sites list.")
    return sites


def safe_filename(value):
    value = re.sub(r'[<>:\"/\\|?*\x00-\x1f]+', "_", value or "").strip(" .")
    return value or "unnamed_site"


def normalize_segment(value):
    value = re.sub(r"^\s*\d+\s*[.)-]?\s*", "", value or "")
    return " ".join(value.split()).casefold()


def classify_file_path(relative_path):
    """Return (department, document_type) when the path matches a target.

    document_type will be a valid Neo4j node label (e.g., RTM, SDD, CONFIG_WORKBOOK, USER_GUIDE).
    """
    segments = [normalize_segment(part) for part in relative_path.split("/") if part]
    for index in range(len(segments)):
        for folder_path, doc_type in DOCUMENT_PATHS:
            end = index + len(folder_path)
            if tuple(segments[index:end]) != folder_path or end >= len(segments):
                continue
            department = TRACKS.get(segments[end])
            if department:
                return department, doc_type
    return None


def collect_document_records(files):
    """Convert collected file metadata into Neo4j-ready document records."""
    records = []
    for item in files:
        classified = classify_file_path(item["path"])
        if not classified:
            continue

        department, document_type = classified
        document_name = item["name"]
        version_match = VERSION_PATTERN.search(document_name)
        extension = Path(document_name).suffix.lstrip(".").lower()

        # Extract version or default to "UNKNOWN" (matches existing Neo4j data)
        version = version_match.group(0).lower() if version_match else "UNKNOWN"

        records.append({
            "department": department,
            "document_name": document_name,
            "document_type": document_type,
            "file_extension": extension,
            "modified_timestamp": item.get("last_modified") or "",
            "sharepoint_path": "/" + item["path"].lstrip("/"),
            "sharepoint_url": item.get("web_url") or "",
            "version": version,
        })

    # Sort for deterministic output
    records.sort(key=lambda row: (
        row["department"].casefold(),
        row["document_type"].casefold(),
        row["sharepoint_path"].casefold(),
    ))
    return records


def scan_folder(graph, drive_id, folder_path, files):
    """Build one nested folder tree and collect file metadata while walking it."""
    children = []
    for item in list_children(graph, drive_id, folder_path):
        name = item.get("name", "(unnamed item)")
        path = "/".join(part for part in (folder_path, name) if part)
        if "folder" in item:
            node = {
                "name": name,
                "type": "folder",
                "children": scan_folder(graph, drive_id, path, files),
            }
        elif "file" in item:
            node = {"name": name, "type": "file"}
            files.append({
                "name": name,
                "path": path,
                "last_modified": item.get("lastModifiedDateTime"),
                "web_url": item.get("webUrl"),
            })
        else:
            node = {"name": name, "type": "item"}
        children.append(node)
    return children


def render_tree_nodes(nodes, prefix="", lines=None):
    if lines is None:
        lines = []
    for index, node in enumerate(nodes):
        is_last = index == len(nodes) - 1
        connector = "└── " if is_last else "├── "
        icon = "📁 " if node["type"] == "folder" else "📄 "
        suffix = "/" if node["type"] == "folder" else ""
        lines.append(f"{prefix}{connector}{icon}{node['name']}{suffix}")
        if node.get("children"):
            render_tree_nodes(
                node["children"],
                prefix + ("    " if is_last else "│   "),
                lines,
            )
    return lines


def export_site(graph, site_entry, output_dir):
    site_url = site_entry.get("url")
    site_name = site_entry.get("name") or site_entry.get("displayName") or "unnamed_site"
    if not site_url:
        raise ValueError(f"Added site {site_name!r} has no URL.")

    parsed_url = urlparse(site_url)
    if not parsed_url.hostname:
        raise ValueError(f"Added site {site_name!r} has an invalid URL: {site_url}")

    site = resolve_site(graph, parsed_url.hostname, parsed_url.path)
    drives = list_drives(graph, site["id"])
    if not drives:
        raise RuntimeError(f"No document libraries found for site {site_name!r}.")

    drive_nodes = []
    files = []
    for drive in drives:
        drive_files = []
        drive_node = {
            "name": drive.get("name") or "Documents",
            "type": "folder",
            "children": scan_folder(graph, drive["id"], "", drive_files),
        }
        drive_nodes.append(drive_node)
        files.extend(drive_files)

    safe_name = safe_filename(site_name)
    site_dir = output_dir / safe_name
    site_dir.mkdir(parents=True, exist_ok=True)

    tree_lines = [f"📁 {site_name}/"]
    render_tree_nodes(drive_nodes, lines=tree_lines)
    (site_dir / "structure.txt").write_text("\n".join(tree_lines) + "\n", encoding="utf-8")

    records = collect_document_records(files)
    json_path = site_dir / f"{safe_name}.json"
    json_path.write_text(
        json.dumps(records, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    return json_path, site_dir / "structure.txt", len(records)


def main():
    parser = argparse.ArgumentParser(
        description="Compare SharePoint sites and export structure and document metadata for additions."
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help=f"Directory for per-site exports (default: {DEFAULT_OUTPUT_DIR})",
    )
    args = parser.parse_args()

    load_dotenv(BASE_DIR / ".env")
    added_sites = run_comparison()
    if not added_sites:
        print("No new SharePoint sites were found.")
        return

    graph = GraphClient(get_access_token())
    output_dir = args.output_dir.expanduser().resolve()
    for site_entry in added_sites:
        json_path, structure_path, record_count = export_site(
            graph,
            site_entry,
            output_dir,
        )
        print(f"Exported {site_entry.get('name') or site_entry.get('displayName')}:")
        print(f"  Structure: {structure_path}")
        print(f"  Documents: {json_path} ({record_count} matching files)")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)