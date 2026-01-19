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
- ASO Bulk Import Tool for Excel-based workspace provisioning
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

#### ASO Bulk Import Tool
Enterprise-grade bulk provisioning from Excel files:
- Place Excel file with prefix `aso_import` in the `bulk/` folder
- Supports both .xlsx and .xls formats
- Multi-step validation process:
  1. Validates required tabs (Webex Users, Webex Side Cars, Webex Auto Attendant, Webex Hunt Groups)
  2. Checks for additional location-specific tabs
  3. Infers and validates location from Webex Users sheet
  4. Validates location outgoing calling permissions (with optional auto-correction)
  5. Validates all data fields (mandatory columns, MAC addresses, phone numbers)
  6. Verifies phone number availability in location
- Preview table before import with confirmation prompt
- Automated workspace provisioning:
  - Creates workspaces with Webex Calling enabled
  - Provisions devices via MAC address
  - Configures call forwarding (no answer, business continuity)
  - Sets custom outgoing calling permissions
- Skips user provisioning (future feature)
- Comprehensive error reporting and status tracking

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

### CSV Bulk Create

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

### ASO Bulk Import (Excel)

Enterprise bulk provisioning tool for large-scale workspace deployments.

#### Excel File Requirements

**File Naming**: Must start with `aso_import` (e.g., `aso_import_site1.xlsx`)

**Required Tabs**:
- Webex Users
- Webex Side Cars
- Webex Auto Attendant
- Webex Hunt Groups
- At least one additional location-specific tab

**Webex Users Sheet Columns** (A-S):

| Column | Name | Required | Description | Valid Values |
|--------|------|----------|-------------|-------------|
| A | - | Optional | - | - |
| B | - | Optional | - | - |
| C | - | Yes | - | - |
| D | Phone Number | Optional | 10-digit phone number | Must be available in location |
| E | Extension | Yes | Extension number | Numeric, validated |
| F | - | Optional | - | - |
| G | - | Optional | - | - |
| H | - | Yes | - | - |
| I | - | Optional | - | - |
| J | User Type | Yes | User or workspace | "user" or "non-user" |
| K | Device Model | Yes | Phone model | Must match PHONE_MODELS or COLLAB_MODELS |
| L | MAC Address | Yes | Device MAC | 12 hex characters (any format) |
| M | Display Name | Yes | Workspace name | Used as workspace displayName |
| N | Forward No Answer | Optional | Forward destination | Phone number for unanswered calls |
| O | Rings Before Forward | Optional | Number of rings | Numeric, max 15 (default: 3) |
| P | - | Optional | - | "yes" or "no" |
| Q | Business Continuity | Optional | Forward destination | Phone number for network disconnect |
| R | - | Optional | - | "yes" or "no" |
| S | Calling Permission | Optional | Permission type | "custom" to apply restrictions |

#### Validation Process

**Step 1: Tab Validation**
- Verifies all required tabs exist
- Confirms at least one additional location tab

**Step 2: Location Validation**
- Infers location from "Location Name" column in Webex Users sheet
- Matches against Webex telephony locations (case-insensitive)
- Displays location ID and calling line ID

**Step 3: Location Outgoing Permissions Check**
- Fetches current location outgoing calling permissions
- Displays permissions table with all call types
- Validates against expected configuration:
  - INTERNAL_CALL: ALLOW with transfer enabled
  - All others: BLOCK with transfer disabled
  - Ignores: CASUAL, URL_DIALING, UNKNOWN
- Prompts to auto-correct if mismatches found
- Continues workflow regardless of user choice

**Step 4: Data Validation**
- Mandatory columns (C, E, H, J, K, L, M) must have values
- MAC addresses: 12 hexadecimal characters, no duplicates
- User Type (J): Must be "user" or "non-user"
- Extension (E): Must be numeric
- Phone Number (D): Must be 10 digits or empty
- Rings (O): Must be numeric and ≤15
- Yes/No fields (P, R): Must be "yes", "no", or empty
- Forward numbers (N, Q): Must be numeric or empty

**Step 5: Phone Number Availability**
- Fetches available PSTN numbers from location
- Filters for unassigned, non-main, ACTIVE numbers
- Validates Column D numbers exist in available pool
- Converts 10-digit to E.164 format (+1XXXXXXXXXX)

#### Import Process

**Preview Phase**:
- Displays table of all items to be imported
- Shows: Row, Type, Name, Extension, Phone, Device
- Counts users vs workspaces
- Requires user confirmation to proceed (default: Yes, press Enter)

**Workspace Creation Phase**:
- Skips rows with User Type = "user" (not yet implemented)
- Creates workspaces with:
  - Display name from Column M
  - Extension from Column E
  - Phone number from Column D (if populated)
  - Location from validated location
  - Webex Calling always enabled
- Provisions device via MAC address (Column L)
- Uses device model from Column K
- Tracks workspace IDs for subsequent configuration

**Call Forwarding Configuration Phase**:
- Configures for each created workspace
- No Answer forwarding (Column N):
  - Destination: Column N value
  - Rings: Column O value (default: 3)
  - Always includes callForwarding structure
- Business Continuity (Column Q):
  - Enabled if Column Q has value
  - Destination: Column Q value

**Outgoing Permissions Configuration Phase**:
- Applies custom permissions if Column S = "custom"
- Configuration:
  - INTERNAL_CALL, TOLL_FREE, NATIONAL: ALLOW with transfer
  - All others: BLOCK without transfer
- Skips if Column S is empty or not "custom"
- Shows status for each workspace

**Side Car Configuration Phase**:
- Prompts user to proceed (default: Yes, press Enter)
- Reads "Webex Side Cars" sheet from Excel
- Extracts target extensions from rows 4-5, column D
- Builds speed dial array from rows 7-34, columns C-D
- Configures device layout with:
  - Custom layout mode
  - 6 line keys (first as PRIMARY_LINE, rest OPEN)
  - KEM_20_KEYS module type
  - Speed dial entries with labels and values
- Applies configuration to all devices matching target extensions
- Shows success/warning for each device

**Hunt Group Configuration Phase**:
- Prompts user to proceed (default: Yes, press Enter)
- Reads "Webex Hunt Groups" sheet from Excel
- Fetches location timezone from Webex API
- Parses hunt groups in 3-row blocks starting at row 4:
  - Column A: Hunt group name
  - Column B: Phone number (optional, skips "N/A")
  - Column C: Extension
  - Column D: Agent extensions (up to 3 per hunt group)
  - Column F: Call policy (default: REGULAR)
  - Column G: Next agent rings (default: 3)
- Maps agent extensions to workspace IDs from created workspaces
- Removes trailing digits from name for customName field
- Displays configuration table for each hunt group
- Asks user to confirm proceeding with each hunt group
- Allows user to modify attributes (name, extension, phoneNumber, policy, nextAgentRings)
- Creates hunt group via POST API with:
  - Extension and phoneNumber as numeric values (not strings)
  - Call policies (policy, waitingEnabled=false, noAnswer settings)
  - Agents array with workspace IDs
  - Hunt group caller ID settings
  - Direct line caller ID with CUSTOM_NAME selection
- Shows success message or error details for each hunt group

**Summary**:
- Total users skipped
- Workspaces created/failed
- Detailed error/warning list
- Note about future user provisioning

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
│   ├── workspaces.csv      # CSV bulk create input
│   ├── workspaces.csv.example  # CSV template
│   └── aso_import*.xlsx    # Excel bulk import files
└── libraries/               # Modular functions
    ├── api_client.py       # API client wrapper
    ├── list_workspaces.py  # List function
    ├── view_workspace.py   # View details function
    ├── create_workspace.py # Create function
    ├── update_workspace.py # Update function
    ├── delete_workspace.py # Delete function
    ├── add_device.py       # Device provisioning
    ├── bulk_create_workspaces.py  # CSV bulk operations
    └── aso_bulk_import.py  # Excel bulk import tool
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
