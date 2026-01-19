# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details

from libraries.list_workspaces import list_workspaces
from libraries.add_device import add_workspace_devices

def configure_workspace_calling(api, workspace_id):
    print("\n--- Configure Workspace Calling ---")
    
    # Get locations
    locations_result = api.call("GET", "locations", params={"orgId": api.org_id})
    if "error" in locations_result:
        print(f"Error fetching locations: {locations_result['error']}")
        return
    
    locations = locations_result.get("items", [])
    if not locations:
        print("No locations found.")
        return
    
    print("\nAvailable Locations:")
    for i, loc in enumerate(locations, 1):
        print(f"{i}. {loc.get('name', 'N/A')} (ID: {loc.get('id', 'N/A')})")
    
    loc_choice = input("Select location number: ").strip()
    try:
        location_id = locations[int(loc_choice) - 1]["id"]
    except (ValueError, IndexError):
        print("Invalid selection.")
        return
    
    extension = input("Enter extension (required): ").strip()
    if not extension:
        print("Extension is required.")
        return
    
    calling_data = {
        "calling": {
            "type": "webexCalling",
            "webexCalling": {
                "extension": extension,
                "locationId": location_id
            }
        }
    }
    
    result = api.call("PUT", f"workspaces/{workspace_id}", data=calling_data)
    
    if "error" in result:
        print(f"Error configuring calling: {result['error']}")
    else:
        print("Calling configured successfully!")

def update_workspace(api):
    print("\n--- Update Workspace ---")
    
    workspaces = list_workspaces(api)
    if not workspaces:
        return
    
    choice = input("\nEnter workspace number to update (or /b to back): ").strip()
    if choice == "/b":
        return
    
    try:
        idx = int(choice) - 1
        workspace = workspaces[idx]
        workspace_id = workspace["id"]
    except (ValueError, IndexError):
        print("Invalid selection.")
        return
    
    print(f"\nUpdating workspace: {workspace.get('displayName')}")
    print("Press Enter to keep current value")
    
    display_name = input(f"Display name [{workspace.get('displayName')}]: ").strip()
    capacity = input(f"Capacity [{workspace.get('capacity', 'N/A')}]: ").strip()
    
    data = {}
    if display_name:
        data["displayName"] = display_name
    if capacity:
        data["capacity"] = int(capacity)
    
    if data:
        result = api.call("PUT", f"workspaces/{workspace_id}", data=data)
        if "error" in result:
            print(f"Error updating workspace: {result['error']}")
        else:
            print("Workspace updated successfully!")
    
    # Ask if user wants to update calling
    update_calling = input("\nUpdate Webex Calling configuration? (y/n): ").strip().lower()
    if update_calling == 'y':
        configure_workspace_calling(api, workspace_id)
    
    # Ask if user wants to manage devices
    manage_devices = input("\nManage devices? (y/n): ").strip().lower()
    if manage_devices == 'y':
        add_workspace_devices(api, workspace_id)
