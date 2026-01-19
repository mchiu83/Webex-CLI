# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details

PHONE_MODELS = [
    "Cisco 6821", "Cisco 6841", "Cisco 6851", "Cisco 6861",
    "Cisco 6861 Wi-Fi", "Cisco 6871", "Cisco 6871",
    "Cisco 7811", "Cisco 7821", "Cisco 7841", "Cisco 7861",
    "Cisco 8811", "Cisco 8841", "Cisco 8851", "Cisco 8861",
    "Cisco 8845", "Cisco 8865", "Cisco 8875",
    "Cisco 9841", "Cisco 9851", "Cisco 9861", "Cisco 9871",
    "Cisco 6823", "Cisco 6825",
    "Cisco 840", "Cisco 860",
    "Cisco 7832", "Cisco 8832",
    "Polycom 5000", "Polycom 6000",
    "Cisco 191", "Cisco 192",
    "Cisco VG400 ATA", "Cisco VG410 ATA", "Cisco VG420 ATA"
]

COLLAB_MODELS = [
    "Cisco Webex Desk Pro", "Cisco Webex Board Pro G2",
    "Cisco Webex Board 55", "Cisco Webex Board 55S", "Cisco Webex Board 70", "Cisco Webex Board 70S", "Cisco Webex Board 85",
    "Cisco Webex Room 55", "Cisco Webex Room 55 Dual", "Cisco Webex Room 70", "Cisco Webex Room 70G2",
    "Cisco Webex Room Kit", "Cisco Webex Room Kit Mini", "Cisco Webex Room Kit Plus",
    "Cisco Webex Room Kit Plus Precision 60", "Cisco Webex Room Kit Pro",
    "Cisco Room Kit EQ", "Cisco Room Kit EQX",
    "Cisco Webex Desk", "Cisco Webex Desk Mini", "Cisco Webex Desk Hub",
    "Cisco Spark Board 55", "Cisco Room Navigator for Table",
    "Cisco WebEx Codec Plus", "CS Codec Pro - stand alone", "Spark Room Kit unit"
]

def add_workspace_devices(api, workspace_id, supported_devices=None):
    print("\n--- Add Devices to Workspace ---")
    
    # If supported_devices not provided, ask user
    if not supported_devices:
        print("\nDevice Type:")
        print("1. Cisco Phones")
        print("2. Collaboration Devices")
        device_type_choice = input("Select device type: ").strip()
        supported_devices = "phones" if device_type_choice == "1" else "collaborationDevices"
    
    models = PHONE_MODELS if supported_devices == "phones" else COLLAB_MODELS
    
    print(f"\nAvailable {'Phone' if supported_devices == 'phones' else 'Collaboration Device'} Models:")
    for i, model in enumerate(models, 1):
        print(f"{i}. {model}")
    
    model_choice = input("\nSelect model number: ").strip()
    try:
        model = models[int(model_choice) - 1]
    except (ValueError, IndexError):
        print("Invalid selection.")
        return
    
    print("\nDevice Provisioning Method:")
    print("1. Activation Code")
    print("2. MAC Address")
    method_choice = input("Select method: ").strip()
    
    if method_choice == "1":
        # Activation Code method
        data = {
            "workspaceId": workspace_id,
            "model": model
        }
        result = api.call("POST", f"devices/activationCode", data=data, params={"orgId": api.org_id})
        
        if "error" in result:
            print(f"Error creating device: {result['error']}")
        else:
            activation_code = result.get("code", "N/A")
            print(f"\nDevice created successfully!")
            print(f"Activation Code: {activation_code}")
            print(f"Device ID: {result.get('id', 'N/A')}")
    
    elif method_choice == "2":
        # MAC Address method
        mac_input = input("\nEnter MAC address: ").strip()
        
        # Parse MAC address - remove all non-alphanumeric characters
        mac_clean = ''.join(c for c in mac_input.upper() if c.isalnum())
        
        if len(mac_clean) != 12:
            print(f"Invalid MAC address. Expected 12 characters, got {len(mac_clean)}.")
            return
        
        # Format as XX:XX:XX:XX:XX:XX
        mac_formatted = ':'.join(mac_clean[i:i+2] for i in range(0, 12, 2))
        
        confirm = input(f"Confirm MAC address: {mac_formatted} (y/n): ").strip().lower()
        if confirm != 'y':
            print("Device creation cancelled.")
            return
        
        data = {
            "mac": mac_formatted,
            "model": model,
            "workspaceId": workspace_id
        }
        result = api.call("POST", "devices", data=data, params={"orgId": api.org_id})
        
        if "error" in result:
            print(f"Error creating device: {result['error']}")
        else:
            print(f"\nDevice created successfully!")
            print(f"Device ID: {result.get('id', 'N/A')}")
    else:
        print("Invalid method selection.")
