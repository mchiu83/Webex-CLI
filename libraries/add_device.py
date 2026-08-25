# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details

PHONE_MODELS = [
    # Cisco MPP Desk Phones
    "Cisco 6821", "Cisco 6841", "Cisco 6851", "Cisco 6861",
    "Cisco 6861 Wi-Fi", "Cisco 6871", "Cisco 6871",
    "Cisco 6823", "Cisco 6825",
    "Cisco 7811", "Cisco 7821", "Cisco 7841", "Cisco 7861",
    "Cisco 8811", "Cisco 8841", "Cisco 8851", "Cisco 8861",
    "Cisco 8845", "Cisco 8865", "Cisco 8875",
    "Cisco 8800 A-KEM", "Cisco 8800 BE-KEM",
    "Cisco 9811", "Cisco 9841", "Cisco 9851", "Cisco 9861", "Cisco 9871",
    # Cisco DECT
    "Cisco 840", "Cisco 860",
    # Cisco Conference Phones
    "Cisco 7832", "Cisco 8832",
    # Cisco ATA
    "Cisco 191", "Cisco 192",
    "Cisco VG400 ATA", "Cisco VG410 ATA", "Cisco VG420 ATA",
    # AudioCodes Phones
    "AudioCodes 425HD", "AudioCodes 445HD", "AudioCodes C450HD",
    # AudioCodes ATA
    "AudioCodes MP-124E (TLS 17FXS)",
    "AudioCodes MP-504", "AudioCodes MP-508", "AudioCodes MP-516",
    "AudioCodes MP-524", "AudioCodes MP-532", "AudioCodes MP-1288",
    "AudioCodes MP202", "AudioCodes MP202R", "AudioCodes MP204", "AudioCodes MP204R",
    # Poly ATA
    "Poly ATA 400", "Poly ATA 402",
    # Poly Conference Phones
    "Poly Trio 8300", "Poly Trio 8500", "Poly Trio 8800", "Poly Trio C60",
    # Poly Desk Phones
    "Poly VVX 101", "Poly VVX 150", "Poly VVX 201", "Poly VVX 250",
    "Poly VVX 301", "Poly VVX 311", "Poly VVX 350",
    "Poly VVX 401", "Poly VVX 411", "Poly VVX 450",
    "Poly VVX 501", "Poly VVX 601",
    # Polycom (legacy branding in Control Hub)
    "Polycom CCX400", "Polycom CCX500", "Polycom CCX505",
    "Polycom CCX600", "Polycom CCX700",
    "Polycom EE100", "Polycom EE220", "Polycom EE300", "Polycom EE320",
    "Polycom EE350", "Polycom EE400", "Polycom EE450",
    "Polycom EE500", "Polycom EE550",
    "Polycom SSIP5000", "Polycom SSIP6000",
    # SNOM Phones
    "SNOM D385",
    "SNOM D713", "SNOM D715", "SNOM D717", "SNOM D735",
    "SNOM D785", "SNOM D787",
    "SNOM D810", "SNOM D812", "SNOM D815",
    # Yealink Desk Phones
    "Yealink T31W", "Yealink T33G", "Yealink T34W", "Yealink T40G",
    "Yealink T41S", "Yealink T42S", "Yealink T43U",
    "Yealink T46S", "Yealink T46U", "Yealink T48S", "Yealink T48U",
    "Yealink T53W", "Yealink T54W", "Yealink T57W",
    "Yealink T58", "Yealink T58V",
    "Yealink T73U", "Yealink T73W", "Yealink T74U", "Yealink T74W", "Yealink T77U",
    "Yealink T85W", "Yealink T87W", "Yealink T88V", "Yealink T88W",
    # Yealink Conference Phones
    "Yealink CP920", "Yealink CP925", "Yealink CP960", "Yealink CP965",
    # Yealink DECT Bases
    "Yealink W52P", "Yealink W56P", "Yealink W60P", "Yealink W70P",
    # Yealink Wi-Fi Handsets
    "Yealink AX83H", "Yealink AX86R",
    # User-friendly aliases (mapped to API strings via DEVICE_MODEL_API_MAP)
    "Yealink AX86",
]

COLLAB_MODELS = [
    "Cisco Webex Desk Pro", "Cisco Webex Board Pro G2",
    "Cisco Webex Board 55", "Cisco Webex Board 55S", "Cisco Webex Board 70", "Cisco Webex Board 70S", "Cisco Webex Board 85",
    "Cisco Webex Room 55", "Cisco Webex Room 55 Dual", "Cisco Webex Room 70", "Cisco Webex Room 70G2",
    "Cisco Webex Room Kit", "Cisco Webex Room Kit Mini", "Cisco Webex Room Kit Plus",
    "Cisco Webex Room Kit Plus Precision 60", "Cisco Webex Room Kit Pro",
    "Cisco Room Kit EQ", "Cisco Room Kit EQX",
    "Cisco Room Bar", "Cisco Room Bar Pro",
    "Cisco Room Navigator", "Cisco Room Navigator for Table",
    "Cisco Desk Camera 4K",
    "Cisco Webex Desk", "Cisco Webex Desk Mini", "Cisco Webex Desk Hub",
    "Cisco Spark Board 55",
    "Cisco WebEx Codec Plus", "CS Codec Pro - stand alone", "Spark Room Kit unit",
    "LG 55UN343H0UA", "GSM LG TV"
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
    
    # Resolve the user-friendly model name to the Webex API model string
    from libraries.workspace_config import DEVICE_MODEL_API_MAP
    api_model = DEVICE_MODEL_API_MAP.get(model, model)

    print("\nDevice Provisioning Method:")
    print("1. Activation Code")
    print("2. MAC Address")
    method_choice = input("Select method: ").strip()
    
    if method_choice == "1":
        # Activation Code method
        data = {
            "workspaceId": workspace_id,
            "model": api_model
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
            "model": api_model,
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
