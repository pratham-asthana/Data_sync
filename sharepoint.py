from pathlib import Path
from dotenv import load_dotenv
import os

SCRIPT_DIR = Path(__file__).resolve().parent
ENV_FILE = SCRIPT_DIR / ".env"

load_dotenv(ENV_FILE, override=True)

SHAREPOINT_URL = os.getenv("SHAREPOINT_URL")

print("Loaded .env:", ENV_FILE)
print("SharePoint URL found:", bool(SHAREPOINT_URL))