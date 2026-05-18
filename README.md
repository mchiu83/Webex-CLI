# Webex Control Hub CLI

A Python CLI tool for bulk provisioning Webex Calling locations via the Webex Control Hub APIs. Designed around an Excel-based import workflow (ASO Bulk Import) that handles the full lifecycle of a store/location setup.

## Features

- **ASO Bulk Import** — full end-to-end provisioning from a single Excel file:
  - Workspace creation with Webex Calling (extension + DID)
  - Device provisioning via MAC address
  - Call forwarding (no answer, business continuity)
  - Outgoing calling permissions
  - Side car / KEM speed dial layout
  - Hunt group creation
  - Auto attendant creation (with audio announcement upload)
  - Call park extension pre-validation
  - Call park group creation
  - Business hours schedule creation (`24-7`, `8-5NBD`)
  - Translation pattern pre-validation
- **Reset Store** — tear down and re-provision an existing location (workspaces, hunt groups, auto attendants, announcements, call park groups/extensions)
- Dual session logging: CLI transcript + raw API call log
- Menu-driven interface with `/b` back navigation

## Requirements

- Python 3.10+
- Windows OS
- Valid Webex API token with full admin permissions
- Organization ID (auto-detected if only one org is available)

## Installation

No dependencies to install. All required libraries (`requests`, `openpyxl`, `xlrd`, and their dependencies) are bundled in the `vendor/` folder.

```bash
python webex.py
```

## Getting Your Webex Bearer Token

The CLI authenticates using a personal access token from the Webex Developer Portal.

1. Go to [https://developer.webex.com](https://developer.webex.com) and sign in with your Webex admin account
2. Click your profile avatar in the top-right corner
3. Your **personal access token** is displayed — click **Copy** to copy it to the clipboard

> This token is valid for **12 hours**. After it expires, return to the developer portal and copy a fresh one.

For long-lived automation, a Webex integration or service account token with a longer TTL is recommended — but for day-to-day use the personal token is sufficient.

## Configuration

Create `credentials.priv` in the project root:

```
token=<your_webex_api_token>
orgid=<your_organization_id>
```

If the file is missing, the CLI will prompt for the token and auto-select the org (or let you pick from a list if multiple exist).


## Usage

```bash
python webex.py
```

### Main Menu

```
1. ASO Bulk Import Tool (All in One)
2. Reset Store
3. Exit
```

---

## ASO Bulk Import Tool

The primary workflow. Reads an Excel file from the `bulk/` folder and provisions a location top-to-bottom.

### Excel File

- Filename must start with `aso_import` (e.g. `aso_import_store0387.xlsx`)
- Place the file in the `bulk/` folder
- Supports `.xlsx` and `.xls`

**Required sheets:**
- `Webex Users`
- `Webex Side Cars`
- `Webex Auto Attendant`
- `Webex Hunt Groups`
- At least one location-specific tab (named after the location)

### Validation Phase

Before any provisioning, the tool runs these checks in order:

| Step | What it checks |
|------|---------------|
| 1 | Required tabs exist; at least one location tab present |
| 2 | Location name from `Webex Users` sheet matches a Webex telephony location |
| 3 | Mandatory columns (C, E, H, J, K, L, M) are populated; MAC/extension/phone formats are valid |
| 4 | Phone numbers in column D exist in the location's available number pool |
| 5 | Translation pattern for the location exists (warns if missing) |
| 6 | Call park extensions match what's defined in the location tab (warns if missing) |
| 7 | Business hours schedules (`24-7`, `8-5NBD`) exist — offers to create them if not |

### Webex Users Sheet Columns

| Col | Field | Required | Notes |
|-----|-------|----------|-------|
| D | Phone Number | Optional | 10-digit DID; must be available in location |
| E | Extension | Yes | Numeric |
| H | Location Name | Yes | Used to resolve the Webex location |
| J | User Type | Yes | `user` (skipped) or `non-user` (workspace) |
| K | Device Model | Yes | Must match a supported Cisco phone model |
| L | MAC Address | Yes | Any format — colons, dashes, plain hex all accepted |
| M | Display Name | Yes | Workspace display name |
| N | Forward No Answer | Optional | Destination phone number |
| O | Rings Before Forward | Optional | Numeric, max 15 (default: 3) |
| P | Forward No Answer Enabled | Optional | `yes` / `no` |
| Q | Business Continuity | Optional | Destination phone number |
| R | Business Continuity Enabled | Optional | `yes` / `no` |
| S | Calling Permission | Optional | `custom` to apply restricted outgoing permissions |

### Provisioning Steps (All in One)

After validation passes, the tool runs each phase in sequence, prompting before each one:

1. **Workspace Import** — creates workspaces, provisions devices, configures call forwarding and outgoing permissions
2. **Side Car Speed Dials** — configures KEM/side car button layouts from the `Webex Side Cars` sheet
3. **Hunt Groups** — creates hunt groups from the `Webex Hunt Groups` sheet; lets you review and edit each one before creation
4. **Auto Attendants** — creates auto attendants from the `Webex Auto Attendant` sheet, uploads WAV announcements, wires menus
5. **Call Park Group** — creates (or updates) the call park group from the location tab, using all non-Paging workspaces as members and all call park extensions as destinations

Each phase can be skipped individually.

### Running Individual Phases

The bootstrap and individual phase runners are also accessible from `webex.py` as standalone methods (not exposed in the main menu, but callable directly for development/debugging):

| Method | What it runs |
|--------|-------------|
| `_run_workspace_import()` | Workspace + device + forwarding + permissions only |
| `_run_sidecar_import()` | Side car speed dials only |
| `_run_huntgroup_import()` | Hunt groups only |
| `_run_call_handler_import()` | Auto attendants only |

---

## Reset Store

Tears down an existing location's provisioned resources and re-provisions from the Excel file. Uses a lighter validation pass (file + location only — skips number availability checks).

Resources deleted/reset (with confirmation at each step):
- Workspaces
- Hunt groups
- Auto attendants
- Location announcements
- Call park groups
- Call park extensions

---

## Auto Attendant Details

The `Webex Auto Attendant` sheet defines up to four AA blocks. For each:

- **Phone number / extension** read from column D of the AA header row
- **Business hours schedule** read from the schedule map (columns J/K, rows 23–29); must be `24-7` or `8-5NBD`
- **Greeting audio** matched by name from the greeting map (column M, rows 5–8); WAV files are uploaded to the location if not already present
- **Key configurations** support: `Transfer to Number`, `Transfer to Sub Menu`, `Repeat Menu`
- **After-hours menu** mirrors the business hours menu

If a phone number is already assigned elsewhere (API error 4201), the tool prompts for a fallback extension rather than failing silently.

---

## Schedule Management

Schedules are validated and auto-created during the ASO bootstrap. Two templates are supported:

| Name | Description |
|------|-------------|
| `24-7` | All 7 days, all day |
| `8-5NBD` | Monday–Friday, 08:00–17:00 |

Schedule names are read from the `Webex Auto Attendant` sheet (column J, rows 23–29). If a required schedule doesn't exist in the location, the tool offers to create it.

---

## Project Structure

```
├── webex.py                          # Entry point and main menu
├── credentials.priv                  # API credentials (gitignored)
├── credentials.priv.example          # Credentials template
├── requirements.txt                  # Reference only — no install needed
├── vendor/                           # Bundled third-party libraries (no pip required)
│   ├── requests/                     # HTTP client
│   ├── urllib3/                      # requests dependency
│   ├── certifi/                      # requests dependency
│   ├── charset_normalizer/           # requests dependency
│   ├── idna/                         # requests dependency
│   ├── openpyxl/                     # .xlsx reader/writer
│   └── xlrd/                         # .xls reader
├── bulk/                             # Import files go here
│   ├── aso_import*.xlsx              # ASO import Excel files
│   └── backup/                       # Archived/previous import files
├── logs/                             # Auto-created per session
│   ├── clisession_YYYYMMDD_HHMMSS.log   # Full CLI transcript
│   └── api_calls_YYYYMMDD_HHMMSS.log    # Raw API requests + responses
└── libraries/
    ├── api_client.py                 # Thin requests wrapper (WebexAPI)
    ├── aso_bulk_import.py            # File selection, Excel reading, workspace import
    ├── aso_validation.py             # All pre-flight validation steps
    ├── workspace_config.py           # Call forwarding, permissions, side car config
    ├── configure_hunt_groups.py      # Hunt group creation
    ├── configure_auto_attendant.py   # Auto attendant creation + audio upload
    ├── configure_call_park_group.py  # Call park group creation/update
    ├── schedule_manager.py           # Schedule validation and creation
    ├── reset_store.py                # Location teardown
    ├── add_device.py                 # Device provisioning helpers
    ├── bulk_create_workspaces.py     # Legacy CSV bulk create (not in main menu)
    ├── create_workspace.py           # Single workspace creation (not in main menu)
    ├── update_workspace.py           # Single workspace update (not in main menu)
    ├── delete_workspace.py           # Single workspace delete (not in main menu)
    ├── list_workspaces.py            # List workspaces (not in main menu)
    └── view_workspace.py             # View workspace details (not in main menu)
```

## Logging

Every session writes two log files to `logs/`:

- `clisession_YYYYMMDD_HHMMSS.log` — everything printed to the terminal
- `api_calls_YYYYMMDD_HHMMSS.log` — every API call with URL, params, request body, response status, and response body

Logs may contain phone numbers, extensions, MAC addresses, and API tokens. Handle accordingly.

## Security Notes

- API tokens grant full admin access — rotate them if exposed
- Log files contain full API responses including sensitive data — restrict access to the `logs/` folder

## License

Copyright (c) 2026 Ming Chiu. Licensed under the MIT License.
