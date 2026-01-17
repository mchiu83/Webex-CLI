# Webex Control Hub CLI

A Python CLI application for managing Webex Control Hub Cloud Calling API, specifically for workspace management.

## Features

- List workspaces
- View detailed workspace information
- Create workspaces with Webex Calling features
- Update workspaces and their configurations
- Delete workspaces
- Add and manage calling devices (Cisco Phones and Collaboration Devices)
- Bulk create workspaces from CSV file
- Session-specific logging for all actions and API calls
- Menu-driven interface with back navigation (/b)
- Modular code structure for easy maintenance

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

## Configuration

### Option 1: Credentials File (Recommended)
Create a file named `credentials.priv` in the same directory as the script:
```
token=<your_webex_api_token>
orgid=<your_organization_id>
```

### Option 2: Manual Entry
If `credentials.priv` is not found:
- The script will prompt you for your API token
- If orgid is not provided, it will fetch and display available organizations for selection

## Usage

Run the application:
```bash
python webex.py
```

### Navigation
- Enter the number corresponding to your choice
- Type `/b` to go back to the previous menu
- Follow on-screen prompts for each operation

### Workspace Operations

#### List Workspaces
Displays all workspaces in your organization with their IDs.

#### View Workspace Details
Shows detailed information including:
- Basic workspace information
- Calling configuration
- Associated devices

#### Create Workspace
Prompts for:
- Display name (required)
- Capacity (optional)
- Type (notSet/focus/huddle/meetingRoom/open/desk/other)
- Supported device type (Cisco Phones or Collaboration Devices)
- Optional: Enable Webex Calling (location, extension, phone number)
- Optional: Add devices via activation code or MAC address

#### Update Workspace
Allows updating:
- Basic workspace properties
- Calling configuration
- Device associations

#### Delete Workspace
Removes a workspace after confirmation.

#### Bulk Create Workspaces
Create multiple workspaces from a CSV file:
- Place `workspaces.csv` in the `bulk/` folder
- CSV columns: id, location, displayName, supportedDevices, type, capacity, calling, extension, phoneNumber, phoneModel, macaddress
- Comprehensive validation before execution
- Preview and confirm before creating
- Detailed results summary

### Device Provisioning

When adding devices to workspaces, you can choose:

**Cisco Phones** (35+ models supported):
- 6800 Series, 7800 Series, 8800 Series, 9800 Series
- IP DECT Series, Conference Phones
- ATAs and VG Gateways

**Collaboration Devices** (25+ models supported):
- Webex Desk Series, Board Series, Room Series
- Room Kits, Codec Plus/Pro

**Provisioning Methods**:
1. **Activation Code**: Generate code for manual device registration
2. **MAC Address**: Direct provisioning with device MAC address
   - Accepts various formats (with/without colons, dashes, spaces)
   - Validates and confirms before creation

## Bulk Operations

### CSV File Format

Place your `workspaces.csv` file in the `bulk/` folder with these columns:

| Column | Required | Description | Valid Values |
|--------|----------|-------------|--------------|
| id | No | Auto-generated after creation | Leave empty |
| location | Optional* | Location name | Must match existing location |
| displayName | Yes | Workspace name | Any string |
| supportedDevices | No | Device type | "phones" or "collaborationDevices" (default) |
| type | No | Workspace type | notSet/focus/huddle/meetingRoom/open/desk/other |
| capacity | No | Room capacity | Number |
| calling | No | Calling type | "none" (default) or "webexCalling" |
| extension | Conditional | Extension number | 4+ digits (required if calling=webexCalling) |
| phoneNumber | No | Phone number | 10 digits |
| phoneModel | No | Device model | Must match PHONE_MODELS or COLLAB_MODELS |
| macaddress | No | Device MAC | 12 alphanumeric characters (no separators) |

*Location is required if calling=webexCalling, otherwise prompted during execution

**Example CSV:**
```csv
id,location,displayName,supportedDevices,type,capacity,calling,extension,phoneNumber,phoneModel,macaddress
,Main Office,Conference Room 1,phones,meetingRoom,10,webexCalling,4001,5551234567,Cisco 8841,
,Main Office,Huddle Space A,collaborationDevices,huddle,4,webexCalling,4002,,Cisco Webex Desk,
,,Open Workspace 1,phones,open,20,none,,,,,
```

### Validation Rules

- **CSV Structure**: Validates headers and field counts
- **displayName**: Mandatory, cannot be empty
- **supportedDevices**: Must be "phones" or "collaborationDevices"
- **calling**: Must be "none" or "webexCalling"
- **extension**: Required for webexCalling, minimum 4 digits, numbers only
- **phoneNumber**: Optional, must be exactly 10 digits
- **phoneModel**: Must match device type (phones vs collaboration devices)
- **macaddress**: Must be 12 alphanumeric characters without separators

All validation errors are reported with specific row numbers before any execution.

## Project Structure

```
Webex-CLI/
├── webex.py                 # Main entry point
├── credentials.priv         # API credentials (not in git)
├── requirements.txt         # Python dependencies
├── logs/                    # Session logs
│   ├── webexapi_*.log      # CLI output transcript
│   └── api_calls_*.log     # API call details
├── bulk/                    # Bulk operation files
│   ├── workspaces.csv      # Bulk create input
│   └── workspaces.csv.example  # Template
└── libraries/               # Modular functions
    ├── api_client.py       # API client wrapper
    ├── list_workspaces.py  # List function
    ├── view_workspace.py   # View details function
    ├── create_workspace.py # Create function
    ├── update_workspace.py # Update function
    ├── delete_workspace.py # Delete function
    ├── add_device.py       # Device provisioning
    └── bulk_create_workspaces.py  # Bulk operations
```

## Logging

All actions and API calls are logged to separate files in the `logs/` folder:
- `webexapi_YYYYMMDD_HHMMSS.log` - Complete CLI output/session transcript
- `api_calls_YYYYMMDD_HHMMSS.log` - All API calls to Webex Control Hub with timestamps, requests, and responses

## API Reference

This application uses the Webex Calling Provisioning APIs:
https://github.com/webex/postman-webex-calling/blob/master/provisioning-api/webex-calling-provisioning-apis.json

## Requirements

- Python 3.6+
- Windows OS
- Valid Webex API token with appropriate permissions
- Organization ID (or will be auto-selected)

## Security Notes

- Keep `credentials.priv` secure and never commit it to version control
- Add `credentials.priv` to `.gitignore`
- API tokens should have minimal required permissions
- Logs may contain sensitive information - handle appropriately
- Bulk CSV files may contain phone numbers and extensions - protect accordingly
