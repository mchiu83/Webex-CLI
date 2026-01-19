# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details

import json
from libraries.list_workspaces import list_workspaces

def view_workspace_details(api, workspace_id=None):
    if not workspace_id:
        workspaces = list_workspaces(api)
        if not workspaces:
            return
        
        choice = input("\nEnter workspace number to view details (or /b to back): ").strip()
        if choice == "/b":
            return
        
        try:
            idx = int(choice) - 1
            workspace_id = workspaces[idx]["id"]
        except (ValueError, IndexError):
            print("Invalid selection.")
            return
    
    print(f"\n--- Workspace Details ---")
    result = api.call("GET", f"workspaces/{workspace_id}")
    
    if "error" in result:
        print(f"Error: {result['error']}")
        return
    
    print(json.dumps(result, indent=2))
    
    # Get calling details
    calling_result = api.call("GET", f"telephony/config/workspaces/{workspace_id}")
    if "error" not in calling_result:
        print("\n--- Calling Configuration ---")
        print(json.dumps(calling_result, indent=2))
    
    # Get devices
    devices_result = api.call("GET", f"workspaces/{workspace_id}/devices")
    if "error" not in devices_result:
        print("\n--- Associated Devices ---")
        print(json.dumps(devices_result, indent=2))
