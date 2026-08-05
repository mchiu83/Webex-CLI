# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details

import re
from libraries.add_device import PHONE_MODELS, COLLAB_MODELS

def _set_device_line_label(api, workspace_id, line_label):
    """
    Set the Line Label on a workspace's device after creation.
    Fetches the calling device ID, then PUTs the members with lineLabel set.
    Returns error string on failure, None on success.
    """
    # Get the calling device ID from the workspace
    devices_result = api.call(
        "GET",
        f"telephony/config/workspaces/{workspace_id}/devices",
        params={"orgId": api.org_id}
    )
    if "error" in devices_result:
        return f"Could not fetch device: {devices_result['error']}"

    devices = devices_result.get('devices', [])
    if not devices:
        return "No device found on workspace after creation"

    device_id = devices[0].get('id')

    # Get current members to preserve existing config
    members_result = api.call(
        "GET",
        f"telephony/config/devices/{device_id}/members",
        params={"orgId": api.org_id}
    )
    if "error" in members_result:
        return f"Could not fetch device members: {members_result['error']}"

    members = members_result.get('members', [])
    if not members:
        # Fallback: construct the member entry for the primary owner
        members = [{
            "id": workspace_id,
            "port": 1,
            "primaryOwner": True,
            "memberType": "PLACE",
            "lineType": "PRIMARY",
            "lineWeight": 1,
            "hotlineEnabled": False,
            "allowCallDeclineEnabled": True,
            "lineLabel": line_label
        }]
    else:
        # Add lineLabel to the primary owner member
        for member in members:
            if member.get('primaryOwner', False) or member.get('id') == workspace_id:
                member['lineLabel'] = line_label
                break
        else:
            # If no match found, set it on the first member
            members[0]['lineLabel'] = line_label

    result = api.call(
        "PUT",
        f"telephony/config/devices/{device_id}/members",
        data={"members": members},
        params={"orgId": api.org_id}
    )

    if "error" in result:
        return result['error']
    return None


def create_workspace_from_row(api, location_data, row, headers, skip_device=False):
    """Create workspace from Excel row data.

    Args:
        skip_device: If True, skip device registration (used for Fax workspace
                     in a Fax/Paging ATA 192 pair — the device lives on the
                     Paging workspace and Fax is added as a shared line after).
    """
    display_name = str(row[12]).strip() if len(row) > 12 else ""
    phone_number = str(row[3]).strip() if len(row) > 3 and row[3] else None
    extension = str(row[4]).strip() if len(row) > 4 else ""
    device_model = str(row[10]).strip() if len(row) > 10 else ""
    mac_address = str(row[11]).strip() if len(row) > 11 else ""
    
    if device_model in PHONE_MODELS:
        supported_devices = "phones"
    elif device_model in COLLAB_MODELS:
        supported_devices = "collaborationDevices"
    else:
        supported_devices = "collaborationDevices"
    
    data = {
        "displayName": display_name,
        "orgId": api.org_id,
        "type": "notSet",
        "supportedDevices": supported_devices,
        "locationId": location_data['id'],
        "calling": {
            "type": "webexCalling",
            "webexCalling": {
                "extension": extension,
                "locationId": location_data['id']
            }
        }
    }
    
    if phone_number and phone_number.isdigit() and len(phone_number) == 10:
        data["calling"]["webexCalling"]["phoneNumber"] = f"+1{phone_number}"

    result = api.call("POST", "workspaces", data=data)
    
    if "error" in result:
        return None, result['error']
    
    workspace_id = result.get("id")
    
    if not skip_device and workspace_id and mac_address and device_model:
        mac_clean = re.sub(r'[-:\s]', '', mac_address).upper()
        mac_formatted = ':'.join(mac_clean[i:i+2] for i in range(0, 12, 2))
        
        device_data = {
            "mac": mac_formatted,
            "model": device_model,
            "workspaceId": workspace_id
        }
        
        device_result = api.call("POST", "devices", data=device_data, params={"orgId": api.org_id})
        
        if "error" in device_result:
            return workspace_id, f"Workspace created but device failed: {device_result['error']}"
        
        # Set Line Label using the Phone Label (On Hook) value from column M
        line_label = display_name  # Column 12 = "Phone Label (On Hook)"
        if line_label:
            label_error = _set_device_line_label(api, workspace_id, line_label)
            if label_error:
                return workspace_id, f"Workspace and device created but line label failed: {label_error}"
    
    return workspace_id, None

def configure_call_forwarding(api, workspace_id, row):
    """Configure call forwarding and business continuity for workspace"""
    forward_no_answer = str(row[13]).strip() if len(row) > 13 and row[13] else None
    num_rings = str(row[14]).strip() if len(row) > 14 and row[14] else "3"
    forward_disconnect = str(row[16]).strip() if len(row) > 16 and row[16] else None
    
    if not forward_no_answer and not forward_disconnect:
        return None
    
    data = {
        "callForwarding": {
            "always": {},
            "busy": {"enabled": False},
            "noAnswer": {
                "enabled": True if forward_no_answer else False,
                "destination": forward_no_answer if forward_no_answer else "",
                "numberOfRings": int(float(num_rings)) if forward_no_answer else 3
            }
        }
    }
    
    if forward_disconnect:
        data["businessContinuity"] = {
            "enabled": True,
            "destination": forward_disconnect
        }
    
    result = api.call("PUT", f"workspaces/{workspace_id}/features/callForwarding", 
                     data=data, params={"orgId": api.org_id})
    
    if "error" in result:
        return result['error']
    
    return None

def configure_outgoing_permission(api, workspace_id, row):
    """Configure outgoing calling permissions for workspace"""
    calling_permission = str(row[18]).strip().lower() if len(row) > 18 and row[18] else None
    
    if calling_permission != 'custom':
        return None, False
    
    data = {
        "useCustomEnabled": True,
        "useCustomPermissions": True,
        "callingPermissions": [
            {"callType": "INTERNAL_CALL", "action": "ALLOW", "transferEnabled": True},
            {"callType": "TOLL_FREE", "action": "ALLOW", "transferEnabled": True},
            {"callType": "NATIONAL", "action": "ALLOW", "transferEnabled": True},
            {"callType": "INTERNATIONAL", "action": "BLOCK", "transferEnabled": False},
            {"callType": "OPERATOR_ASSISTED", "action": "BLOCK", "transferEnabled": False},
            {"callType": "CHARGEABLE_DIRECTORY_ASSISTED", "action": "BLOCK", "transferEnabled": False},
            {"callType": "SPECIAL_SERVICES_I", "action": "BLOCK", "transferEnabled": False},
            {"callType": "SPECIAL_SERVICES_II", "action": "BLOCK", "transferEnabled": False},
            {"callType": "PREMIUM_SERVICES_I", "action": "BLOCK", "transferEnabled": False},
            {"callType": "PREMIUM_SERVICES_II", "action": "BLOCK", "transferEnabled": False}
        ]
    }
    
    result = api.call("PUT", f"workspaces/{workspace_id}/features/outgoingPermission", 
                     data=data, params={"orgId": api.org_id})
    
    if "error" in result:
        return result['error'], True
    
    return None, True

def configure_ata_fax_line(api, paging_device_id, paging_workspace_id, fax_workspace_id):
    """
    Assign the Fax workspace as FXS port 2 on the Paging workspace's ATA 192 device.
    Uses PUT /telephony/config/devices/{deviceId}/members with the full field set
    required by the API (derived from GET response inspection).

    Args:
        paging_device_id:    Calling device ID of the ATA 192 (from the Paging workspace)
        paging_workspace_id: Workspace ID of the Paging workspace (port 1, primaryOwner)
        fax_workspace_id:    Workspace ID of the Fax workspace (port 2)
    Returns:
        error string on failure, None on success
    """
    data = {
        "members": [
            {
                "id": paging_workspace_id,
                "port": 1,
                "primaryOwner": True,
                "memberType": "PLACE",
                "lineType": "PRIMARY",
                "lineWeight": 1,
                "hotlineEnabled": False,
                "allowCallDeclineEnabled": True,
                "t38FaxCompressionEnabled": False
            },
            {
                "id": fax_workspace_id,
                "port": 2,
                "primaryOwner": False,
                "memberType": "PLACE",
                "lineType": "PRIMARY",
                "lineWeight": 1,
                "hotlineEnabled": False,
                "allowCallDeclineEnabled": False,
                "t38FaxCompressionEnabled": False
            }
        ]
    }

    result = api.call(
        "PUT",
        f"telephony/config/devices/{paging_device_id}/members",
        data=data,
        params={"orgId": api.org_id}
    )

    if "error" in result:
        return result['error']
    return None


def configure_side_car_speed_dials(api, workspace_map, data_rows, filepath, read_excel_sheet):
    """Configure side car speed dials for devices"""
    print(f"\n{'='*60}")
    print("Configuring Side Car Speed Dials")
    print(f"{'='*60}")
    
    sidecar_data = read_excel_sheet(filepath, 'Webex Side Cars')
    if not sidecar_data or len(sidecar_data) < 7:
        print("  Skipped: No side car data found")
        return
    
    target_extensions = []
    for row_idx in [3, 4]:
        if len(sidecar_data) > row_idx and len(sidecar_data[row_idx]) > 3:
            ext = str(sidecar_data[row_idx][3]).strip() if sidecar_data[row_idx][3] else None
            if ext:
                target_extensions.append(ext)
    
    if not target_extensions:
        print("  Skipped: No target extensions found in rows 4-5")
        return
    
    ext_to_workspace = {}
    for row_idx, workspace_id in workspace_map.items():
        row = data_rows[row_idx - 2]
        extension = str(row[4]).strip() if len(row) > 4 else ""
        if extension in target_extensions:
            ext_to_workspace[extension] = workspace_id
    
    device_map = {}
    for extension, workspace_id in ext_to_workspace.items():
        print(f"\n  Fetching device for extension {extension}...")
        devices_result = api.call("GET", f"telephony/config/workspaces/{workspace_id}/devices",
                                 params={"orgId": api.org_id})
        
        if "error" in devices_result:
            print(f"    Warning: Failed to fetch devices - {devices_result['error']}")
            continue
        
        devices = devices_result.get('devices', [])
        if devices:
            device_id = devices[0].get('id')
            device_map[extension] = device_id
            print(f"    Found device ID: {device_id}")
        else:
            print(f"    Warning: No devices found")
    
    if not device_map:
        print("\n  Skipped: No devices found for configuration")
        return
    
    kem_keys = []
    for row_idx in range(6, 34):
        if len(sidecar_data) <= row_idx:
            break
        
        row = sidecar_data[row_idx]
        label = str(row[2]).strip() if len(row) > 2 and row[2] else None
        value = str(row[3]).strip() if len(row) > 3 and row[3] else None
        
        if label and value:
            kem_keys.append({
                "kemModuleIndex": 1,
                "kemKeyIndex": len(kem_keys) + 1,
                "kemKeyType": "SPEED_DIAL",
                "kemKeyLabel": label,
                "kemKeyValue": value
            })
    
    if not kem_keys:
        print("\n  Skipped: No speed dials found in rows 7-34")
        return
    
    print(f"\n  Found {len(kem_keys)} speed dial entries")
    
    for extension, device_id in device_map.items():
        print(f"\n  Configuring device for extension {extension}...")
        
        layout_data = {
            "layoutMode": "CUSTOM",
            "userReorderEnabled": False,
            "lineKeys": [
                {"lineKeyIndex": 1, "lineKeyType": "PRIMARY_LINE"},
                {"lineKeyIndex": 2, "lineKeyType": "OPEN"},
                {"lineKeyIndex": 3, "lineKeyType": "OPEN"},
                {"lineKeyIndex": 4, "lineKeyType": "OPEN"},
                {"lineKeyIndex": 5, "lineKeyType": "OPEN"},
                {"lineKeyIndex": 6, "lineKeyType": "OPEN"}
            ],
            "kemModuleType": "KEM_20_KEYS",
            "kemKeys": kem_keys
        }
        
        result = api.call("PUT", f"telephony/config/devices/{device_id}/layout",
                         data=layout_data, params={"orgId": api.org_id})
        
        if "error" in result:
            print(f"    Warning: Failed to configure - {result['error']}")
        else:
            print(f"    Success: Side car speed dials configured")
