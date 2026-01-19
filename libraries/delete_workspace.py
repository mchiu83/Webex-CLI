# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details

from libraries.list_workspaces import list_workspaces

def delete_workspace(api):
    print("\n--- Delete Workspace ---")
    
    workspaces = list_workspaces(api)
    if not workspaces:
        return
    
    choice = input("\nEnter workspace number to delete (or /b to back): ").strip()
    if choice == "/b":
        return
    
    try:
        idx = int(choice) - 1
        workspace = workspaces[idx]
        workspace_id = workspace["id"]
    except (ValueError, IndexError):
        print("Invalid selection.")
        return
    
    confirm = input(f"Are you sure you want to delete '{workspace.get('displayName')}'? (yes/no): ").strip().lower()
    if confirm != "yes":
        print("Deletion cancelled.")
        return
    
    result = api.call("DELETE", f"workspaces/{workspace_id}")
    
    if "error" in result:
        print(f"Error deleting workspace: {result['error']}")
    else:
        print("Workspace deleted successfully!")
