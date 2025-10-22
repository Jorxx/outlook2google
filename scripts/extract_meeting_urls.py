#!/usr/bin/env python3
"""
Process Microsoft 365 Calendar Events JSON and extract meeting URLs to CSV
Creates a CSV file for each user with event IDs and meeting URLs
"""

import json
import csv
import argparse
import os
from datetime import datetime

def extract_meeting_urls_from_json(json_file, output_dir="."):
    """Extract meeting URLs from JSON and create CSV files per user"""
    
    print(f"📖 Reading JSON file: {json_file}")
    
    with open(json_file, "r", encoding="utf-8") as f:
        data = json.load(f)
    
    events = data.get("events", [])
    print(f"📊 Found {len(events)} events")
    
    # Group events by user
    user_events = {}
    for event in events:
        user_email = event.get("user_email", "unknown")
        if user_email not in user_events:
            user_events[user_email] = []
        user_events[user_email].append(event)
    
    print(f"👥 Found events for {len(user_events)} users")
    
    # Process each user
    for user_email, user_event_list in user_events.items():
        # Extract username from email (part before @)
        username = user_email.split("@")[0]
        csv_filename = os.path.join(output_dir, f"{username}_events.csv")
        
        print(f"📝 Processing {len(user_event_list)} events for {user_email}")
        
        with open(csv_filename, "w", newline="", encoding="utf-8") as csvfile:
            fieldnames = [
                "event_id",
                "event_name", 
                "event_description",
                "start_date",
                "end_date",
                "timezone",
                "meeting_url",
                "location",
                "attendees_emails",
                "attendees_count"
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for event in user_event_list:
                # Extract meeting URL from location field
                location = event.get("location", "")
                meeting_url = ""
                
                # Check if location contains a URL
                if location and ("http" in location.lower() or "zoom.us" in location.lower() or "teams.microsoft.com" in location.lower()):
                    meeting_url = location
                
                # Also check the dedicated meeting_url field
                if not meeting_url:
                    meeting_url = event.get("meeting_url", "")
                
                # Extract attendees emails
                attendees = event.get("attendees", [])
                attendees_emails = "; ".join([attendee.get("email", "") for attendee in attendees if attendee.get("email")])
                attendees_count = len(attendees) if attendees else 0
                
                # Write row
                writer.writerow({
                    "event_id": event.get("event_id", ""),
                    "event_name": event.get("event_name", ""),
                    "event_description": event.get("event_description", ""),
                    "start_date": event.get("start_date", ""),
                    "end_date": event.get("end_date", ""),
                    "timezone": event.get("timezone", ""),
                    "meeting_url": meeting_url,
                    "location": location,
                    "attendees_emails": attendees_emails,
                    "attendees_count": attendees_count
                })
        
        print(f"✅ Created {csv_filename}")
    
    print(f"\n🎉 Processing complete!")
    print(f"📁 Created CSV files for {len(user_events)} users in {output_dir}")

def main():
    parser = argparse.ArgumentParser(description="Extract meeting URLs from MS events JSON to CSV")
    parser.add_argument("--input", type=str, required=True, 
                       help="Input JSON file (e.g., theo_events.json)")
    parser.add_argument("--output-dir", type=str, default=".", 
                       help="Output directory for CSV files (default: current directory)")
    parser.add_argument("--user", type=str, default=None,
                       help="Process only this specific user (email)")
    args = parser.parse_args()
    
    if not os.path.exists(args.input):
        print(f"❌ Input file not found: {args.input}")
        return
    
    # Create output directory if it doesn't exist
    os.makedirs(args.output_dir, exist_ok=True)
    
    extract_meeting_urls_from_json(args.input, args.output_dir)

if __name__ == "__main__":
    main()
