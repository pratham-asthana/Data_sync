
import os
import json
import requests
import msal

from pathlib import Path
from urllib.parse import urlparse, quote
from dotenv import load_dotenv

# Load .env from the script's directory
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

SITE_URL = os.getenv(
    "SHAREPOINT_SITE_URL",
    "https://orbrick.sharepoint.com/"
)

GRAPH_URL = "https://graph.microsoft.com/v1.0"
TARGET_HOST = urlparse(SITE_URL).hostname


def get_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )

    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    if "access_token" not in result:
        raise RuntimeError(
            result.get("error_description", str(result))
        )

    return result["access_token"]


def get_all_pages(url, headers):
    results = []

    while url:
        response = requests.get(
            url, headers=headers, timeout=60
        )

        if not response.ok:
            raise RuntimeError(
                f"Graph API error {response.status_code}: "
                f"{response.text}"
            )

        data = response.json()
        results.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    return results


def main():
    if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
        raise ValueError(
            "Check TENANT_ID, CLIENT_ID and CLIENT_SECRET in .env"
        )

    headers = {
        "Authorization": f"Bearer {get_token()}",
        "Accept": "application/json",
    }

    # Resolve the tenant root site
    root_url = (
        f"{GRAPH_URL}/sites/"
        f"{TARGET_HOST}:/"
    )

    response = requests.get(
        root_url, headers=headers, timeout=60
    )
    response.raise_for_status()
    root_site = response.json()

    # Search tenant-wide. An empty search returns broad results,
    # but Graph search is not guaranteed to enumerate every site.
    search_url = f"{GRAPH_URL}/sites?search={quote('*')}"
    sites = get_all_pages(search_url, headers)

    # Include the tenant root explicitly
    sites.append(root_site)

    # Keep unique sites hosted on the specified tenant hostname
    unique = {}

    for site in sites:
        web_url = site.get("webUrl", "")
        hostname = urlparse(web_url).hostname

        if hostname and hostname.lower() == TARGET_HOST.lower():
            unique[site["id"]] = {
                "name": site.get("name"),
                "displayName": site.get("displayName"),
                "url": web_url,
                "id": site.get("id"),
            }

    output = sorted(
        unique.values(),
        key=lambda item: (item["url"] or "").lower()
    )

    output_path = BASE_DIR / "sharepoint_sites.json"

    with open(output_path, "w", encoding="utf-8") as file:
        json.dump(output, file, indent=4, ensure_ascii=False)

    print(f"Tenant: {TARGET_HOST}")
    print(f"Sites discovered: {len(output)}")

    for site in output:
        print(f"{site['displayName'] or site['name']}: {site['url']}")

    print(f"\nSaved to: {output_path}")


if __name__ == "__main__":
    main()
