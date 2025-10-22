#!/bin/bash

# Calendar Migration Tool - Outlook to Google Calendar
# This script migrates calendar events from Microsoft 365 to Google Workspace
# Usage: ./outlook2google.sh [--dry-run] [--config CONFIG_FILE]

set -e  # Exit on any error

# Default configuration file
CONFIG_FILE="${CONFIG_FILE:-config/migration.conf}"

# Check for command line options
DRY_RUN=false
while [[ $# -gt 0 ]]; do
    case $1 in
        --dry-run)
            DRY_RUN=true
            shift
            ;;
        --config)
            CONFIG_FILE="$2"
            shift 2
            ;;
        -h|--help)
            echo "Usage: $0 [--dry-run] [--config CONFIG_FILE]"
            echo ""
            echo "Options:"
            echo "  --dry-run    Show what would be executed without creating events"
            echo "  --config     Specify configuration file (default: config/migration.conf)"
            echo "  -h, --help   Show this help message"
            exit 0
            ;;
        *)
            echo "Unknown option: $1"
            echo "Use --help for usage information"
            exit 1
            ;;
    esac
done

# Load configuration
if [[ ! -f "$CONFIG_FILE" ]]; then
    echo "ERROR: Configuration file '$CONFIG_FILE' not found!"
    echo "Please create a configuration file or specify one with --config"
    exit 1
fi

source "$CONFIG_FILE"

# Validate required configuration
if [[ -z "$GAM_PATH" ]]; then
    echo "ERROR: GAM_PATH not set in configuration file"
    exit 1
fi

if [[ -z "$PYTHON_PATH" ]]; then
    echo "ERROR: PYTHON_PATH not set in configuration file"
    exit 1
fi

if [[ -z "$SCRIPT_DIR" ]]; then
    echo "ERROR: SCRIPT_DIR not set in configuration file"
    exit 1
fi

# Set up paths
GAM_CMD="$GAM_PATH"
PYTHON_CMD="$PYTHON_PATH"
EVENT_EXTRACTION_SCRIPT="$SCRIPT_DIR/ms_events_export_json.py"
JSON2CSV_SCRIPT="$SCRIPT_DIR/extract_meeting_urls.py"

# Check if dry-run mode
if [[ "$DRY_RUN" == "true" ]]; then
    echo "🔍 DRY RUN MODE - No events will be created"
    echo ""
fi

echo "Enter the username to transfer the calendar from:"
read FROM_USER

boats_email=$(echo $FROM_USER | awk -F'@' '{print $1}')"@boats.com"

echo "================================================"
echo "Starting the migration from $FROM_USER to $boats_email"
echo "================================================"

# Enable the python environment if specified
if [[ -n "$PYTHON_ENV" ]]; then
    echo "Activating Python environment: $PYTHON_ENV"
    source "$PYTHON_ENV"
fi

# Extract events from Microsoft 365
echo "Extracting events from Microsoft 365..."
$PYTHON_CMD $EVENT_EXTRACTION_SCRIPT --user $FROM_USER --output ${FROM_USER}_events.json --pretty

# Convert JSON to CSV
echo "Converting events to CSV format..."
$PYTHON_CMD $JSON2CSV_SCRIPT --input ${FROM_USER}_events.json --output ${FROM_USER}_events.csv

echo "================================================"
echo "Creating calendar events in Google Calendar"
echo "================================================"

# Function to convert ISO date to GAM format
convert_date_for_gam() {
    local iso_date=$1
    # Convert from 2025-10-23T08:00:00.0000000 to 2025-10-23 08:00:00
    echo "$iso_date" | sed 's/T/ /' | sed 's/\.[0-9]*$//'
}

# Function to create calendar event using GAM
create_calendar_event() {
    local event_name="$1"
    local event_description="$2"
    local start_date="$3"
    local end_date="$4"
    local meeting_url="$5"
    local attendees="$6"
    
    # Convert dates for GAM
    local gam_start=$(convert_date_for_gam "$start_date")
    local gam_end=$(convert_date_for_gam "$end_date")
    
    # Prepare attendees list (convert semicolon-separated to comma-separated)
    local gam_attendees=""
    if [[ -n "$attendees" && "$attendees" != "null" ]]; then
        gam_attendees=$(echo "$attendees" | tr ';' ',' | sed 's/,$//')
    fi
    
    # Prepare description with meeting URL if available
    local full_description="$event_description"
    if [[ -n "$meeting_url" && "$meeting_url" != "null" && "$meeting_url" != "" ]]; then
        full_description="${event_description}\n\nMeeting URL: ${meeting_url}"
    fi
    
    echo "Creating event: $event_name"
    echo "Start: $gam_start"
    echo "End: $gam_end"
    echo "Attendees: $gam_attendees"
    echo ""
    
    if [[ "$DRY_RUN" == "true" ]]; then
        echo "🔍 DRY RUN - Would execute:"
        if [[ -n "$gam_attendees" ]]; then
            echo "gam user $boats_email add calendar event \"$event_name\" start \"$gam_start\" end \"$gam_end\" description \"$full_description\" attendees \"$gam_attendees\""
        else
            echo "gam user $boats_email add calendar event \"$event_name\" start \"$gam_start\" end \"$gam_end\" description \"$full_description\""
        fi
        echo "✅ DRY RUN - Event would be created: $event_name"
    else
        # Create the calendar event using GAM
        if [[ -n "$gam_attendees" ]]; then
            $GAM_CMD user $boats_email add calendar event "$event_name" start "$gam_start" end "$gam_end" description "$full_description" attendees "$gam_attendees"
        else
            $GAM_CMD user $boats_email add calendar event "$event_name" start "$gam_start" end "$gam_end" description "$full_description"
        fi
        
        if [[ $? -eq 0 ]]; then
            echo "✅ Event created successfully: $event_name"
        else
            echo "❌ Failed to create event: $event_name"
        fi
    fi
    echo ""
}

# Check if CSV file exists
if [[ ! -f "${FROM_USER}_events.csv" ]]; then
    echo "ERROR: CSV file ${FROM_USER}_events.csv not found!"
    exit 1
fi

# Read CSV file and create events (skip header)
echo "Reading events from: ${FROM_USER}_events.csv"
line_count=0
created_count=0
failed_count=0

while IFS=',' read -r user_email event_id event_name event_description start_date end_date meeting_url attendees is_cancelled created_date modified_date raw_event; do
    ((line_count++))
    
    # Skip header row
    if [[ $line_count -eq 1 ]]; then
        continue
    fi
    
    # Skip cancelled events
    if [[ "$is_cancelled" == "True" ]]; then
        echo "Skipping cancelled event: $event_name"
        continue
    fi
    
    # Skip events with empty names
    if [[ -z "$event_name" || "$event_name" == "null" ]]; then
        echo "Skipping event with empty name"
        continue
    fi
    
    # Create the calendar event
    create_calendar_event "$event_name" "$event_description" "$start_date" "$end_date" "$meeting_url" "$attendees"
    
    # In dry-run mode, always count as successful
    if [[ "$DRY_RUN" == "true" ]]; then
        ((created_count++))
    else
        if [[ $? -eq 0 ]]; then
            ((created_count++))
        else
            ((failed_count++))
        fi
    fi
    
done < "${FROM_USER}_events.csv"

echo "================================================"
if [[ "$DRY_RUN" == "true" ]]; then
    echo "Calendar Migration DRY RUN Summary:"
    echo "Total events processed: $((line_count - 1))"
    echo "Events that would be created: $created_count"
    echo "Events that would fail: $failed_count"
    echo ""
    echo "🔍 This was a dry run - no actual events were created"
else
    echo "Calendar Migration Summary:"
    echo "Total events processed: $((line_count - 1))"
    echo "Events created successfully: $created_count"
    echo "Events failed: $failed_count"
fi
echo "================================================"

# Show user info
echo "User information:"
$GAM_CMD user $boats_email
