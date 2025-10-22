# Calendar Migration Tool

A comprehensive tool for migrating calendar events from Microsoft 365 (Outlook) to Google Workspace (Google Calendar) using GAM (Google Apps Manager).

## Features

- 🔄 **Full Migration**: Extract events from Microsoft 365 and create them in Google Calendar
- 🔍 **Dry-Run Mode**: Preview what will be migrated without making changes
- 📅 **Event Details**: Preserves event titles, descriptions, dates, times, and attendees
- 🔗 **Meeting URLs**: Includes Zoom, Teams, and other meeting links in event descriptions
- ⚙️ **Configurable**: Easy configuration through config files
- 📊 **Progress Tracking**: Detailed logging and summary reports
- 🚫 **Smart Filtering**: Automatically skips cancelled events

## Prerequisites

### Required Software

1. **GAM (Google Apps Manager)**
   - Download from: https://github.com/GAM-team/GAM
   - Must be configured with appropriate Google Workspace permissions

2. **Python 3.7+**
   - Required for Microsoft Graph API integration

3. **Microsoft Graph API Access**
   - Azure AD application with Calendar.Read permissions
   - Client ID, Client Secret, and Tenant ID

### Required Permissions

- **Google Workspace**: Calendar creation permissions for target users
- **Microsoft 365**: Calendar read permissions for source users

## Installation

1. **Clone or download this repository**
   ```bash
   git clone <your-repo-url>
   cd calendar-migration-tool
   ```

2. **Set up configuration**
   ```bash
   cp config/migration.conf.example config/migration.conf
   # Edit config/migration.conf with your paths
   ```

3. **Install Python dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Configure Microsoft Graph API**
   - Update the Python scripts with your Azure AD credentials
   - Ensure proper permissions are granted

## Configuration

### Main Configuration (`config/migration.conf`)

```bash
# GAM Configuration
GAM_PATH="/path/to/your/gam"

# Python Configuration  
PYTHON_PATH="/path/to/your/python3"
PYTHON_ENV="/path/to/your/venv/bin/activate"  # Optional

# Script Directory
SCRIPT_DIR="$(dirname "$0")/scripts"
```

### Microsoft Graph API Configuration

Update the Python scripts with your Azure AD application details:

```python
# In ms_events_export_json.py
CLIENT_ID = "your-client-id"
CLIENT_SECRET = "your-client-secret" 
TENANT_ID = "your-tenant-id"
```

## Usage

### Basic Migration

```bash
./outlook2google.sh
```

The script will prompt you for:
- Source user email (Microsoft 365)
- Target user email (Google Workspace)

### Dry-Run Mode

Preview what will be migrated without making changes:

```bash
./outlook2google.sh --dry-run
```

### Custom Configuration

```bash
./outlook2google.sh --config /path/to/custom/config.conf
```

### Help

```bash
./outlook2google.sh --help
```

## Process Overview

1. **Extract Events**: Uses Microsoft Graph API to export calendar events
2. **Convert Format**: Transforms JSON data to CSV for processing
3. **Create Events**: Uses GAM to create events in Google Calendar
4. **Report Results**: Provides detailed summary of migration

## File Structure

```
calendar-migration-tool/
├── outlook2google.sh          # Main migration script
├── scripts/
│   ├── ms_events_export_json.py    # Microsoft Graph API integration
│   └── extract_meeting_urls.py     # JSON to CSV converter
├── config/
│   └── migration.conf.example      # Configuration template
├── docs/
│   └── README.md                   # This file
└── requirements.txt                # Python dependencies
```

## Event Data Preserved

- ✅ Event title and description
- ✅ Start and end dates/times
- ✅ Attendee lists
- ✅ Meeting URLs (Zoom, Teams, etc.)
- ✅ Event locations
- ❌ Recurring event patterns (converted to individual events)
- ❌ Cancelled events (automatically skipped)

## Troubleshooting

### Common Issues

1. **GAM not found**
   - Verify GAM_PATH in configuration file
   - Ensure GAM is properly installed and configured

2. **Python script errors**
   - Check PYTHON_PATH in configuration
   - Verify all dependencies are installed
   - Check Microsoft Graph API credentials

3. **Permission errors**
   - Verify Google Workspace permissions
   - Check Microsoft Graph API permissions
   - Ensure proper authentication

### Debug Mode

For detailed logging, check the generated files:
- `{username}_events.json` - Raw Microsoft Graph data
- `{username}_events.csv` - Processed event data

## Security Notes

- ⚠️ **No sensitive data is hardcoded** in the scripts
- 🔐 Store credentials securely (environment variables recommended)
- 🚫 Never commit configuration files with real credentials
- 🔒 Use least-privilege permissions for all accounts

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review the configuration examples
3. Open an issue in the repository

## Changelog

### Version 1.0.0
- Initial release
- Full calendar migration functionality
- Dry-run mode
- Configurable paths and settings
- Comprehensive documentation
