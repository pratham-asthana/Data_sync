import argparse
import json
import sys
import subprocess
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
LIST_SCRIPT = BASE_DIR / "list_sharepoint_sites.py"
JSON_FILE = BASE_DIR / "sharepoint_sites.json"


def load_sites():
    if not JSON_FILE.exists():
        raise FileNotFoundError(
            f"Previous site list not found: {JSON_FILE}"
        )

    with open(JSON_FILE, "r", encoding="utf-8") as file:
        data = json.load(file)

    if not isinstance(data, list):
        raise ValueError("Expected a JSON list of SharePoint sites.")

    return data


def save_sites(sites):
    with open(JSON_FILE, "w", encoding="utf-8") as file:
        json.dump(sites, file, indent=4, ensure_ascii=False)


def main():
    parser = argparse.ArgumentParser(description="Compare SharePoint sites and detect additions.")
    parser.add_argument(
        "--json",
        action="store_true",
        help="Output result as JSON"
    )
    args = parser.parse_args()
    
    if not LIST_SCRIPT.exists():
        raise FileNotFoundError(
            f"Site-listing script not found: {LIST_SCRIPT}"
        )

    # Preserve the previous snapshot before refreshing it.
    previous_sites = load_sites()

    previous_ids = {
        site.get("id")
        for site in previous_sites
        if site.get("id")
    }

    previous_urls = {
        site.get("url", "").rstrip("/").lower()
        for site in previous_sites
        if site.get("url")
    }

    # Run the existing script to fetch the latest sites.
    result = subprocess.run(
        [sys.executable, str(LIST_SCRIPT)],
        cwd=BASE_DIR,
        capture_output=True,
        text=True,
    )

    if result.returncode != 0:
        print("Failed to retrieve the latest SharePoint sites.")
        if result.stderr:
            print(result.stderr)
        raise SystemExit(result.returncode)

    # The listing script refreshes the JSON file.
    current_sites = load_sites()

    # Identify sites that were not in the previous snapshot.
    added_sites = []
    known_ids = set(previous_ids)
    known_urls = set(previous_urls)

    for site in current_sites:
        site_id = site.get("id")
        url = site.get("url", "").rstrip("/").lower()

        if (
            (site_id and site_id not in known_ids)
            or (url and url not in known_urls)
        ):
            added_sites.append({
                "name": site.get("displayName") or site.get("name"),
                "url": site.get("url"),
                "id": site_id,
                "displayName": site.get("displayName"),
            })

            if site_id:
                known_ids.add(site_id)
            if url:
                known_urls.add(url)

    if not added_sites:
        if args.json:
            print(json.dumps({"added_sites": []}))
        else:
            print("False")
        return

    # Merge old and new sites, preserving existing entries.
    merged_sites = list(previous_sites)
    merged_ids = set(previous_ids)
    merged_urls = set(previous_urls)

    for site in current_sites:
        site_id = site.get("id")
        url = site.get("url", "").rstrip("/").lower()

        if (
            (site_id and site_id not in merged_ids)
            or (url and url not in merged_urls)
        ):
            merged_sites.append(site)

            if site_id:
                merged_ids.add(site_id)
            if url:
                merged_urls.add(url)

    save_sites(merged_sites)

    if args.json:
        print(json.dumps({"added_sites": added_sites}))
    else:
        print("True")
        print("Added sites:")
        for site in added_sites:
            print(f"Name: {site['name']}")
            print(f"URL: {site['url']}")
            print("-" * 40)

        print(f"Updated JSON: {JSON_FILE}")


if __name__ == "__main__":
    main()