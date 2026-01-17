#!/usr/bin/env python3
import requests
import json
import os
import sys
import logging
from datetime import datetime
from typing import Optional, Dict, List
import msvcrt
import threading

class TeeOutput:
    def __init__(self, *files):
        self.files = files
    
    def write(self, data):
        for f in self.files:
            f.write(data)
            f.flush()
    
    def flush(self):
        for f in self.files:
            f.flush()

class WebexCLI:
    def __init__(self):
        self.token = None
        self.org_id = None
        self.base_url = "https://webexapis.com/v1"
        self.session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.setup_logging()
        self.load_credentials()
        
    def setup_logging(self):
        os.makedirs("logs", exist_ok=True)
        
        # CLI output logger - capture all console output
        cli_log = f"logs/webexapi_{self.session_id}.log"
        self.cli_log_file = open(cli_log, 'w')
        sys.stdout = TeeOutput(sys.__stdout__, self.cli_log_file)
        sys.stderr = TeeOutput(sys.__stderr__, self.cli_log_file)
        
        # API calls logger
        api_log = f"logs/api_calls_{self.session_id}.log"
        self.api_logger = logging.getLogger("webex_api")
        self.api_logger.setLevel(logging.INFO)
        api_handler = logging.FileHandler(api_log)
        api_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
        self.api_logger.addHandler(api_handler)
        self.api_logger.propagate = False
        
        print(f"Session started: {self.session_id}")
        
    def load_credentials(self):
        if os.path.exists("credentials.priv"):
            try:
                with open("credentials.priv", "r") as f:
                    for line in f:
                        line = line.strip()
                        if line.startswith("token="):
                            self.token = line.split("=", 1)[1]
                        elif line.startswith("orgid="):
                            self.org_id = line.split("=", 1)[1]
                print("Credentials loaded from credentials.priv")
            except Exception as e:
                print(f"Error loading credentials: {e}")
        
        if not self.token:
            self.token = input("Enter Webex API Token: ").strip()
        
        if not self.org_id:
            orgs_result = self.api_call("GET", "organizations")
            if "error" in orgs_result:
                print(f"Error fetching organizations: {orgs_result['error']}")
                self.org_id = input("Enter Organization ID: ").strip()
            else:
                orgs = orgs_result.get("items", [])
                if not orgs:
                    print("No organizations found.")
                    self.org_id = input("Enter Organization ID: ").strip()
                elif len(orgs) == 1:
                    self.org_id = orgs[0]["id"]
                    print(f"Using organization: {orgs[0].get('displayName', 'N/A')}")
                else:
                    print("\nAvailable Organizations:")
                    for i, org in enumerate(orgs, 1):
                        print(f"{i}. {org.get('displayName', 'N/A')} (ID: {org.get('id', 'N/A')})")
                    org_choice = input("Select organization number: ").strip()
                    try:
                        self.org_id = orgs[int(org_choice) - 1]["id"]
                    except (ValueError, IndexError):
                        print("Invalid selection.")
                        self.org_id = input("Enter Organization ID: ").strip()
            
    def api_call(self, method: str, endpoint: str, data: Optional[Dict] = None, params: Optional[Dict] = None) -> Dict:
        url = f"{self.base_url}/{endpoint}"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        self.api_logger.info(f"API Call: {method} {url}")
        if params:
            self.api_logger.info(f"Params: {json.dumps(params)}")
        if data:
            self.api_logger.info(f"Data: {json.dumps(data)}")
            
        try:
            response = requests.request(method, url, headers=headers, json=data, params=params)
            self.api_logger.info(f"Response Status: {response.status_code}")
            self.api_logger.info(f"Response: {response.text}")
            
            if response.status_code in [200, 201, 204]:
                return response.json() if response.text else {}
            else:
                self.api_logger.error(f"API Error: {response.status_code} - {response.text}")
                return {"error": response.text, "status_code": response.status_code}
        except Exception as e:
            self.api_logger.error(f"Exception during API call: {e}")
            return {"error": str(e)}
    
    def check_back_key(self):
        if msvcrt.kbhit():
            key = msvcrt.getch()
            if key == b'\x00' or key == b'\xe0':
                key = msvcrt.getch()
                if key == b'0':  # Alt+B
                    return True
        return False
    
    def display_menu(self, title: str, options: List[str]) -> str:
        print(f"\n{'='*60}")
        print(f"{title}")
        print(f"{'='*60}")
        for i, option in enumerate(options, 1):
            print(f"{i}. {option}")
        print("/b. Back")
        print(f"{'='*60}")
        
        choice = input("Enter choice: ").strip()
        return choice
    
    def list_workspaces(self):
        print("\n--- List Workspaces ---")
        params = {"orgId": self.org_id}
        result = self.api_call("GET", "workspaces", params=params)
        
        if "error" in result:
            print(f"Error: {result['error']}")
            return None
        
        workspaces = result.get("items", [])
        if not workspaces:
            print("No workspaces found.")
            return None
            
        print(f"\nFound {len(workspaces)} workspace(s):")
        for i, ws in enumerate(workspaces, 1):
            print(f"{i}. {ws.get('displayName', 'N/A')} (ID: {ws.get('id', 'N/A')})")
        
        return workspaces
    
    def view_workspace_details(self, workspace_id: str = None):
        if not workspace_id:
            workspaces = self.list_workspaces()
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
        result = self.api_call("GET", f"workspaces/{workspace_id}")
        
        if "error" in result:
            print(f"Error: {result['error']}")
            return
        
        print(json.dumps(result, indent=2))
        
        # Get calling details
        calling_result = self.api_call("GET", f"telephony/config/workspaces/{workspace_id}")
        if "error" not in calling_result:
            print("\n--- Calling Configuration ---")
            print(json.dumps(calling_result, indent=2))
        
        # Get devices
        devices_result = self.api_call("GET", f"workspaces/{workspace_id}/devices")
        if "error" not in devices_result:
            print("\n--- Associated Devices ---")
            print(json.dumps(devices_result, indent=2))
    
    def create_workspace(self):
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
            "orgId": self.org_id,
            "type": workspace_type,
            "supportedDevices": supported_devices
        }
        
        if capacity:
            data["capacity"] = int(capacity)
        
        # Ask if user wants to enable Webex Calling
        enable_calling = input("\nEnable Webex Calling for this workspace? (y/n): ").strip().lower()
        if enable_calling == 'y':
            # Get locations
            locations_result = self.api_call("GET", "locations", params={"orgId": self.org_id})
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
        
        result = self.api_call("POST", "workspaces", data=data)
        
        if "error" in result:
            print(f"Error creating workspace: {result['error']}")
            return
        
        workspace_id = result.get("id")
        print(f"Workspace created successfully! ID: {workspace_id}")
        
        # Ask if user wants to add devices
        add_devices = input("\nAdd devices to this workspace? (y/n): ").strip().lower()
        if add_devices == 'y':
            self.add_workspace_devices(workspace_id, supported_devices)
    
    def configure_workspace_calling(self, workspace_id: str):
        print("\n--- Configure Workspace Calling ---")
        
        # Get locations
        locations_result = self.api_call("GET", "locations", params={"orgId": self.org_id})
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
        
        if extension:
            calling_data["calling"]["webexCalling"]["extension"] = extension
        
        result = self.api_call("PUT", f"workspaces/{workspace_id}", data=calling_data)
        
        if "error" in result:
            print(f"Error configuring calling: {result['error']}")
        else:
            print("Calling configured successfully!")
    
    def add_workspace_devices(self, workspace_id: str, supported_devices: str = None):
        print("\n--- Add Devices to Workspace ---")
        
        # If supported_devices not provided, ask user
        if not supported_devices:
            print("\nDevice Type:")
            print("1. Cisco Phones")
            print("2. Collaboration Devices")
            device_type_choice = input("Select device type: ").strip()
            supported_devices = "phones" if device_type_choice == "1" else "collaborationDevices"
        
        # Define device models
        phone_models = [
            "Cisco IP Phone 6821", "Cisco IP Phone 6841", "Cisco IP Phone 6851", "Cisco IP Phone 6861",
            "Cisco 6861", "Cisco IP Phone 6871", "Cisco IP Phone 6871 with color display",
            "Cisco 7811", "Cisco 7821", "Cisco 7841", "Cisco IP Phone 7861",
            "Cisco 8811", "Cisco 8841", "Cisco 8851", "Cisco 8861",
            "Cisco 8845", "Cisco 8865", "Cisco 8875",
            "Cisco 9841", "Cisco 9851", "Cisco 9861", "Cisco 9871",
            "Cisco IP DECT 6800 Series", "Cisco IP DECT 6823 Handset", "Cisco IP DECT 6825 Handset",
            "Cisco IP DECT 6825 Handset Ruggedized", "Cisco IP DECT DBS 110 Base Station", "Cisco IP DECT 210 Base-Station",
            "Cisco 840", "Cisco 860",
            "Cisco IP Conference Phone 7832", "Cisco IP Conference Phone 8832",
            "Polycom SoundStation IP 5000", "Polycom SoundStation IP 6000",
            "ATA191 Multiplatform", "ATA192 Multiplatform",
            "VG400 ATA", "VG410 ATA", "VG420 ATA"
        ]
        
        collab_models = [
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
        
        models = phone_models if supported_devices == "phones" else collab_models
        
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
            result = self.api_call("POST", f"devices/activationCode", data=data, params={"orgId": self.org_id})
            
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
            result = self.api_call("POST", "devices", data=data, params={"orgId": self.org_id})
            
            if "error" in result:
                print(f"Error creating device: {result['error']}")
            else:
                print(f"\nDevice created successfully!")
                print(f"Device ID: {result.get('id', 'N/A')}")
        else:
            print("Invalid method selection.")
    
    def update_workspace(self):
        print("\n--- Update Workspace ---")
        
        workspaces = self.list_workspaces()
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
            result = self.api_call("PUT", f"workspaces/{workspace_id}", data=data)
            if "error" in result:
                print(f"Error updating workspace: {result['error']}")
            else:
                print("Workspace updated successfully!")
        
        # Ask if user wants to update calling
        update_calling = input("\nUpdate Webex Calling configuration? (y/n): ").strip().lower()
        if update_calling == 'y':
            self.configure_workspace_calling(workspace_id)
        
        # Ask if user wants to manage devices
        manage_devices = input("\nManage devices? (y/n): ").strip().lower()
        if manage_devices == 'y':
            self.add_workspace_devices(workspace_id)
    
    def delete_workspace(self):
        print("\n--- Delete Workspace ---")
        
        workspaces = self.list_workspaces()
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
        
        result = self.api_call("DELETE", f"workspaces/{workspace_id}")
        
        if "error" in result:
            print(f"Error deleting workspace: {result['error']}")
        else:
            print("Workspace deleted successfully!")
    
    def workspace_menu(self):
        while True:
            choice = self.display_menu(
                "Workspace Management",
                [
                    "List Workspaces",
                    "View Workspace Details",
                    "Create Workspace",
                    "Update Workspace",
                    "Delete Workspace"
                ]
            )
            
            if choice == "/b":
                break
            elif choice == "1":
                self.list_workspaces()
                input("\nPress Enter to continue...")
            elif choice == "2":
                self.view_workspace_details()
                input("\nPress Enter to continue...")
            elif choice == "3":
                self.create_workspace()
                input("\nPress Enter to continue...")
            elif choice == "4":
                self.update_workspace()
                input("\nPress Enter to continue...")
            elif choice == "5":
                self.delete_workspace()
                input("\nPress Enter to continue...")
            else:
                print("Invalid choice. Please try again.")
    
    def main_menu(self):
        print(f"\nWelcome to Webex Control Hub CLI")
        print(f"Organization ID: {self.org_id}")
        print(f"Session ID: {self.session_id}")
        
        while True:
            choice = self.display_menu(
                "Main Menu",
                [
                    "Workspace Management",
                    "Exit"
                ]
            )
            
            if choice == "/b" or choice == "2":
                print("\nExiting...")
                print("Session ended")
                self.cli_log_file.close()
                break
            elif choice == "1":
                self.workspace_menu()
            else:
                print("Invalid choice. Please try again.")

def main():
    try:
        cli = WebexCLI()
        cli.main_menu()
    except KeyboardInterrupt:
        print("\n\nInterrupted by user. Exiting...")
        sys.exit(0)
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
