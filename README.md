# Webex Control Hub CLI

A Python CLI application for managing Webex Control Hub Cloud Calling API, specifically for workspace management.

## Features

- List workspaces
- View detailed workspace information
- Create workspaces with Webex Calling features
- Update workspaces and their configurations
- Delete workspaces
- Add and manage calling devices
- Session-specific logging for all actions and API calls
- Menu-driven interface with back navigation (/b)

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
If `credentials.priv` is not found, the script will prompt you for credentials.

## Usage

Run the application:
```bash
python webex_cli.py
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
- Optional: Configure Webex Calling (location, extension)
- Optional: Add devices

#### Update Workspace
Allows updating:
- Basic workspace properties
- Calling configuration
- Device associations

#### Delete Workspace
Removes a workspace after confirmation.

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
- Organization ID

## Security Notes

- Keep `credentials.priv` secure and never commit it to version control
- Add `credentials.priv` to `.gitignore`
- API tokens should have minimal required permissions
- Logs may contain sensitive information - handle appropriately
