from libraries.add_device import add_workspace_devices

def create_workspace(api):
    print("\n--- Create Workspace ---")
    
    display_name = input("Enter workspace display name: ").strip()
    if not display_name:
        print("Display name is required.")
        return
    
    capacity = input("Enter capacity (optional, press Enter to skip): ").strip()
    workspace_type = input("Enter type (notSet/focus/huddle/meetingRoom/open/desk/other) [default: notSet]: ").strip() or "notSet"
    
    print("\nSupported Device Type:")
    print("1. Cisco Phones")
    print("2. Collaboration Devices")
    device_choice = input("Select device type: ").strip()
    
    if device_choice == "1":
        supported_devices = "phones"
    elif device_choice == "2":
        supported_devices = "collaborationDevices"
    else:
        print("Invalid selection. Defaulting to phones.")
        supported_devices = "phones"
    
    data = {
        "displayName": display_name,
        "orgId": api.org_id,
        "type": workspace_type,
        "supportedDevices": supported_devices
    }
    
    if capacity:
        data["capacity"] = int(capacity)
    
    # Ask if user wants to enable Webex Calling
    enable_calling = input("\nEnable Webex Calling for this workspace? (y/n): ").strip().lower()
    if enable_calling == 'y':
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
        
        extension = input("Enter extension: ").strip()
        if not extension:
            print("Extension is required.")
            return
        
        phone_number = input("Enter phone number (optional, press Enter to skip): ").strip()
        
        data["locationId"] = location_id
        data["calling"] = {
            "type": "webexCalling",
            "webexCalling": {
                "extension": extension,
                "locationId": location_id
            }
        }
        
        if phone_number:
            data["calling"]["webexCalling"]["phoneNumber"] = phone_number
    
    result = api.call("POST", "workspaces", data=data)
    
    if "error" in result:
        print(f"Error creating workspace: {result['error']}")
        return
    
    workspace_id = result.get("id")
    print(f"Workspace created successfully! ID: {workspace_id}")
    
    # Ask if user wants to add devices
    add_devices = input("\nAdd devices to this workspace? (y/n): ").strip().lower()
    if add_devices == 'y':
        add_workspace_devices(api, workspace_id, supported_devices)
