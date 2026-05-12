import os
import json
import requests
from dotenv import load_dotenv

load_dotenv()

TOKEN = os.environ.get("META_TOKEN")
APP_ID = os.environ.get("META_APP_ID")
APP_SECRET = os.environ.get("META_APP_SECRET")

GRAPH_BASE = "https://graph.facebook.com/v19.0"


def safe_get(url, params):
    try:
        return requests.get(url, params=params, timeout=20)
    except requests.RequestException as exc:
        print(f"Request failed for {url}: {exc}")
        return None


def print_json(title, payload):
    print(f"\n=== {title} ===")
    print(json.dumps(payload, indent=2))


def run_preflight(token):
    print("=== Preflight Checks ===")

    me_resp = safe_get(f"{GRAPH_BASE}/me", {"access_token": token, "fields": "id,name"})
    if me_resp is not None:
        print(f"/me status: {me_resp.status_code}")
        print_json("/me response", me_resp.json())

    perms_resp = safe_get(f"{GRAPH_BASE}/me/permissions", {"access_token": token})
    if perms_resp is not None:
        print(f"/me/permissions status: {perms_resp.status_code}")
        perms_data = perms_resp.json()
        print_json("/me/permissions response", perms_data)

        granted = {
            p.get("permission")
            for p in perms_data.get("data", [])
            if p.get("status") == "granted"
        }
        print(f"granted scopes: {sorted(granted)}")
        print(f"ads_read granted: {'ads_read' in granted}")

    if APP_ID and APP_SECRET:
        app_token = f"{APP_ID}|{APP_SECRET}"
        debug_resp = safe_get(
            f"{GRAPH_BASE}/debug_token",
            {"input_token": token, "access_token": app_token},
        )
        if debug_resp is not None:
            print(f"/debug_token status: {debug_resp.status_code}")
            print_json("/debug_token response", debug_resp.json())
    else:
        print("Skipping /debug_token (set META_APP_ID and META_APP_SECRET to enable it).")

if not TOKEN:
    raise ValueError("META_TOKEN environment variable is not set. Run: export META_TOKEN='your_token'")

run_preflight(TOKEN)

url = f"{GRAPH_BASE}/ads_archive"
params = {
    "access_token": TOKEN,
    "search_terms": "Coach clearance",
    "ad_reached_countries": '["US"]',
    "ad_type": "ALL",
    "fields": "id,ad_creative_body,page_name,page_id,page_created_time,ad_delivery_start_time,ad_delivery_stop_time,impressions,spend,advertiser_domains",
    "limit": 10
}

response = safe_get(url, params)
if response is None:
    raise SystemExit(1)

data = response.json()
print(f"\n=== /ads_archive status: {response.status_code} ===")

if "error" in data:
    print("API Error:")
    print(json.dumps(data["error"], indent=2))
else:
    print(json.dumps(data, indent=2))