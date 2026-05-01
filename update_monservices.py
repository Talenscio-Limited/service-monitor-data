"""
update_monservices.py
Updates the MonServices SharePoint list from status.json using Azure AD app credentials.
Runs as a step in the GitHub Actions status-poll workflow.
 
Requires environment variables:
  SHAREPOINT_TENANT_ID     - Azure AD Directory (tenant) ID
  SHAREPOINT_CLIENT_ID     - Azure AD Application (client) ID
  SHAREPOINT_CLIENT_SECRET - Azure AD client secret
"""
 
import json
import os
import sys
from datetime import datetime, timezone
 
import requests
 
# ── Configuration ──────────────────────────────────────────────────────────────
 
TENANT_ID     = os.environ["SHAREPOINT_TENANT_ID"]
CLIENT_ID     = os.environ["SHAREPOINT_CLIENT_ID"]
CLIENT_SECRET = os.environ["SHAREPOINT_CLIENT_SECRET"]
 
SHAREPOINT_SITE = "talensciolimited.sharepoint.com"
SITE_PATH       = "/sites/ProductManagement"
LIST_NAME       = "MonServices"
 
STATUS_JSON_PATH = "status.json"
 
# Map status.json status values to MonServices Choice values
STATUS_MAP = {
    "operational":  "Operational",
    "degraded":     "Degraded",
    "incident":     "Incident",
    "maintenance":  "Maintenance",
}
 
# Map status.json service names to MonServices ServiceName values
# Must match exactly what is stored in the ServiceName column
SERVICE_NAME_MAP = {
    "Talenscio Platform": "Talenscio Platform",
    "ANS Hosting":        "ANS Hosting",
    "Daily.co":           "Daily.co",
    "TalkJS":             "TalkJS",
    "Mailgun / Sinch":    "Mailgun / Sinch",
    "Miro":               "Miro",
    "Stripe":             "Stripe",
    "TawkTo":             "TawkTo",
}
 
 
# ── Step 1: Get Azure AD access token ─────────────────────────────────────────
 
def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type":    "client_credentials",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         f"https://{SHAREPOINT_SITE}/.default",
    }
    resp = requests.post(url, data=data, timeout=30)
    resp.raise_for_status()
    token = resp.json().get("access_token")
    if not token:
        raise ValueError(f"No access token returned: {resp.json()}")
    print("✓ Access token obtained")
    return token
 
 
# ── Step 2: Get all MonServices list items ─────────────────────────────────────
 
def get_list_items(token):
    site_url = f"https://{SHAREPOINT_SITE}{SITE_PATH}"
    url = (
        f"https://graph.microsoft.com/v1.0/sites/"
        f"{SHAREPOINT_SITE}:{SITE_PATH}:/lists/{LIST_NAME}/items"
        f"?expand=fields(select=id,ServiceName,CurrentStatus,StatusDescription,"
        f"LastChecked,ActiveIncidentCount)"
    )
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    items = resp.json().get("value", [])
    print(f"✓ Retrieved {len(items)} MonServices rows")
    return items
 
 
# ── Step 3: Update a single list item ─────────────────────────────────────────
 
def update_list_item(token, item_id, fields):
    url = (
        f"https://graph.microsoft.com/v1.0/sites/"
        f"{SHAREPOINT_SITE}:{SITE_PATH}:/lists/{LIST_NAME}/items/{item_id}/fields"
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type":  "application/json",
    }
    resp = requests.patch(url, headers=headers, json=fields, timeout=30)
    resp.raise_for_status()
 
 
# ── Step 4: Load status.json ───────────────────────────────────────────────────
 
def load_status_json():
    with open(STATUS_JSON_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    services = data.get("services", [])
    print(f"✓ Loaded status.json — {len(services)} services")
    return services
 
 
# ── Main ───────────────────────────────────────────────────────────────────────
 
def main():
    # Load status data
    services = load_status_json()
 
    # Build lookup: service name → status data
    status_lookup = {}
    for svc in services:
        name = svc.get("service", "")
        status_lookup[name] = svc
 
    # Authenticate
    token = get_access_token()
 
    # Get current SharePoint list items
    items = get_list_items(token)
 
    # Build lookup: ServiceName → item id
    sp_lookup = {}
    for item in items:
        fields = item.get("fields", {})
        svc_name = fields.get("ServiceName", "")
        sp_lookup[svc_name] = item.get("id")
 
    # Update each service
    updated = 0
    errors  = 0
    now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
 
    for json_name, sp_name in SERVICE_NAME_MAP.items():
        svc_data = status_lookup.get(json_name)
        item_id  = sp_lookup.get(sp_name)
 
        if not svc_data:
            print(f"⚠ No status data found for '{json_name}' — skipping")
            continue
        if not item_id:
            print(f"⚠ No MonServices row found for '{sp_name}' — skipping")
            continue
 
        raw_status   = svc_data.get("status", "unknown").lower()
        mapped_status = STATUS_MAP.get(raw_status, "Unknown")
        description   = svc_data.get("description", "")
        last_checked  = svc_data.get("lastChecked", now_iso)
 
        # Count active incidents — latestIssue present and not "None"
        latest_issue = svc_data.get("latestIssue", "None")
        active_count = 0 if (not latest_issue or latest_issue == "None") else 1
 
        fields_to_update = {
            "CurrentStatus":       mapped_status,
            "StatusDescription":   description,
            "LastChecked":         last_checked,
            "ActiveIncidentCount": active_count,
        }
 
        try:
            update_list_item(token, item_id, fields_to_update)
            print(f"✓ {sp_name}: {mapped_status}")
            updated += 1
        except Exception as exc:
            print(f"✗ {sp_name}: update failed — {exc}")
            errors += 1
 
    print(f"\nDone — {updated} updated, {errors} errors")
    if errors:
        sys.exit(1)
 
 
if __name__ == "__main__":
    main()
