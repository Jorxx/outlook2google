#!/usr/bin/env python3
"""
Microsoft 365 Calendar Events Export to JSON
Exports calendar events from Microsoft Graph to JSON format
"""

import os
import json
import argparse
from datetime import datetime
from msal import ConfidentialClientApplication
import requests

# ==============================
# CONFIGURATION
# ==============================
KEYS_FILE = "keys.txt"

def _load_dotenv(dotenv_path: str = ".env") -> None:
    """Minimal .env loader without external deps."""
    if not os.path.exists(dotenv_path):
        return
    try:
        with open(dotenv_path, "r", encoding="utf-8") as fh:
            for raw in fh:
                line = raw.strip()
                if not line or line.startswith("#"):
                    continue
                if line.startswith("export "):
                    line = line[len("export "):]
                if "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip().strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = value
    except Exception:
        pass

# Load environment variables early
_load_dotenv()

def _load_keys_from_file(path: str):
    tenant_id = os.environ.get("TENANT_ID")
    client_id = os.environ.get("CLIENT_ID")
    client_secret = os.environ.get("CLIENT_SECRET")

    if not os.path.exists(path):
        return tenant_id, client_id, client_secret

    with open(path, "r", encoding="utf-8") as fh:
        for raw in fh:
            line = raw.strip()
            if not line:
                continue
            norm = line.lower().replace(" ", "")
            if norm.startswith("tenantid"):
                tenant_id = line.split(":", 1)[-1].split("=", 1)[-1].strip()
            elif norm.startswith("appclientid") or norm.startswith("clientid"):
                client_id = line.split(":", 1)[-1].split("=", 1)[-1].strip()
            elif norm.startswith("value") or norm.startswith("clientsecret"):
                client_secret = line.split(":", 1)[-1].split("=", 1)[-1].strip()

    return tenant_id, client_id, client_secret

TENANT_ID, CLIENT_ID, CLIENT_SECRET = _load_keys_from_file(KEYS_FILE)

# ==============================
# AUTHENTICATION
# ==============================
def get_ms_graph_token():
    if not TENANT_ID or not CLIENT_ID or not CLIENT_SECRET:
        raise RuntimeError("Missing TENANT_ID/CLIENT_ID/CLIENT_SECRET. Populate keys.txt or environment variables.")

    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if not token or "access_token" not in token:
        error_desc = token.get("error_description") if isinstance(token, dict) else str(token)
        raise RuntimeError(f"Failed to get MS Graph token. {error_desc}")
    return token["access_token"]

# ==============================
# MICROSOFT GRAPH FUNCTIONS
# ==============================
def get_ms_events(token, user_email):
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/calendar/events?$top=1000"
    headers = {"Authorization": f"Bearer {token}"}
    events = []
    while url:
        resp = requests.get(url, headers=headers)
        data = resp.json()
        events.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return events

def extract_event_data(event, user_email):
    """Extract relevant fields from Microsoft Graph event"""
    
    # Event name
    event_name = event.get("subject", "No Title")
    
    # Event description
    description = event.get("bodyPreview", "")
    
    # Start and end dates
    start_time = event.get("start", {}).get("dateTime", "")
    end_time = event.get("end", {}).get("dateTime", "")
    timezone = event.get("start", {}).get("timeZone", "UTC")
    
    # Meeting URL
    meeting_url = ""
    online_meeting = event.get("onlineMeeting", {})
    if online_meeting and online_meeting.get("joinUrl"):
        meeting_url = online_meeting["joinUrl"]
    
    # Event attendees
    attendees = []
    for attendee in event.get("attendees", []):
        if "emailAddress" in attendee:
            attendees.append({
                "email": attendee["emailAddress"]["address"],
                "name": attendee["emailAddress"].get("name", ""),
                "response": attendee.get("status", {}).get("response", "none")
            })
    
    # Additional useful fields
    event_id = event.get("id", "")
    is_cancelled = event.get("isCancelled", False)
    created_date = event.get("createdDateTime", "")
    modified_date = event.get("lastModifiedDateTime", "")
    location = event.get("location", {}).get("displayName", "")
    
    return {
        "user_email": user_email,
        "event_id": event_id,
        "event_name": event_name,
        "event_description": description,
        "start_date": start_time,
        "end_date": end_time,
        "timezone": timezone,
        "meeting_url": meeting_url,
        "attendees": attendees,
        "location": location,
        "is_cancelled": is_cancelled,
        "created_date": created_date,
        "modified_date": modified_date,
        "online_meeting": online_meeting,  # Full online meeting data
        "raw_event": event  # Full raw data for reference
    }

def export_user_events(user_email, token, debug=False):
    """Export events for a single user to JSON"""
    print(f"📅 Exporting calendar for {user_email}")
    
    try:
        ms_events = get_ms_events(token, user_email)
        print(f"   Found {len(ms_events)} events")
        
        exported_events = []
        exported_count = 0
        
        for event in ms_events:
            # Skip cancelled events unless debug mode
            if event.get("isCancelled") and not debug:
                continue
                
            event_data = extract_event_data(event, user_email)
            exported_events.append(event_data)
            exported_count += 1
            
            if debug:
                print(f"   📝 Exported: {event_data['event_name']}")
                if event_data['meeting_url']:
                    print(f"      🔗 Meeting URL: {event_data['meeting_url']}")
        
        print(f"   ✅ Exported {exported_count} events")
        return exported_events
        
    except Exception as e:
        print(f"   ❌ Error exporting {user_email}: {e}")
        return []

# ==============================
# MAIN FUNCTION
# ==============================
def main():
    parser = argparse.ArgumentParser(description="Export Microsoft 365 calendar events to JSON")
    parser.add_argument("--output", type=str, default="ms_events_export.json", 
                       help="Output JSON file (default: ms_events_export.json)")
    parser.add_argument("--user", type=str, required=True, 
                       help="User email to export calendar from")
    parser.add_argument("--debug", action="store_true", help="Include cancelled events and debug info")
    parser.add_argument("--pretty", action="store_true", help="Pretty print JSON with indentation")
    args = parser.parse_args()
    
    # Get Microsoft Graph token
    print("🔐 Authenticating with Microsoft Graph...")
    token = get_ms_graph_token()
    print("✅ Authentication successful")
    
    # Export events
    exported_events = export_user_events(args.user, token, debug=args.debug)
    
    # Create output data structure
    output_data = {
        "export_info": {
            "user_email": args.user,
            "export_date": datetime.now().isoformat(),
            "total_events": len(exported_events),
            "debug_mode": args.debug
        },
        "events": exported_events
    }
    
    # Write to JSON file
    with open(args.output, "w", encoding="utf-8") as f:
        if args.pretty:
            json.dump(output_data, f, indent=2, ensure_ascii=False)
        else:
            json.dump(output_data, f, ensure_ascii=False)
    
    print(f"\n🎉 Export complete!")
    print(f"📊 Total events exported: {len(exported_events)}")
    print(f"📁 Output file: {args.output}")
    
    # Show summary
    events_with_meetings = sum(1 for event in exported_events if event.get("meeting_url"))
    print(f"🔗 Events with meeting URLs: {events_with_meetings}")
    
    print(f"\n💡 Next steps:")
    print(f"   - Review the JSON file to verify data quality")
    print(f"   - Use this data to create Google Calendar events")
    print(f"   - Run with --debug to see more details")
    print(f"   - Run with --pretty for readable JSON formatting")

if __name__ == "__main__":
    main()
