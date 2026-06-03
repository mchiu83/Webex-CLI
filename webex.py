#!/usr/bin/env python3
# Copyright (c) 2026 Ming Chiu
# Licensed under the MIT License - see LICENSE file for details

import os
import sys

# ---------------------------------------------------------------------------
# Vendor path — add bundled third-party packages so no pip install is needed
# ---------------------------------------------------------------------------
_vendor_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vendor")
if _vendor_path not in sys.path:
    sys.path.insert(0, _vendor_path)

import logging
from datetime import datetime
from typing import List

from libraries.api_client import WebexAPI
from libraries.aso_bulk_import import aso_bulk_import_tool, select_import_file, read_excel_sheet, process_bulk_import

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
        self.session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.setup_logging()
        self.load_credentials()
        self.api = WebexAPI(self.token, self.org_id, self.api_logger)
        
    def setup_logging(self):
        os.makedirs("logs", exist_ok=True)
        
        # CLI output logger - capture all console output
        cli_log = f"logs/clisession_{self.session_id}.log"
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
            temp_api = WebexAPI(self.token, None, self.api_logger)
            orgs_result = temp_api.call("GET", "organizations")
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
    
    def display_menu(self, title: str, options: List[str], show_back: bool = True) -> str:
        print(f"\n{'='*60}")
        print(f"{title}")
        print(f"{'='*60}")
        for i, option in enumerate(options, 1):
            print(f"{i}. {option}")
        if show_back:
            print("/b. Back")
        print(f"{'='*60}")
        choice = input("Enter choice: ").strip()
        return choice

    # ------------------------------------------------------------------
    # Shared bootstrap: locate file + run validations + resolve location
    # Returns (filepath, location_data, additional_tabs) or None on failure
    # ------------------------------------------------------------------
    def _aso_bootstrap(self):
        from libraries.aso_validation import (
            validate_excel_file,
            validate_location,
            validate_webex_users_data,
            validate_available_numbers,
            validate_translation_pattern,
            validate_call_park_extensions
        )
        from libraries.schedule_manager import validate_and_create_schedules

        if not os.path.exists('bulk'):
            print("Status: FAILED - 'bulk' folder not found")
            os.makedirs('bulk')
            print("Please place your 'aso_import' Excel file in the 'bulk' folder and try again.")
            return None

        print("\nSearching for valid import files in bulk folder...")
        filepath = select_import_file()

        if not filepath:
            print("Status: FAILED - No valid import file selected.")
            return None

        print(f"Status: PASS - Using file: {filepath}")

        is_valid, additional_tabs = validate_excel_file(filepath)
        if not is_valid:
            print("\nValidation failed. Please fix the issues and try again.")
            return None

        if additional_tabs:
            print(f"\nAdditional tabs detected and cached:")
            for i, tab in enumerate(additional_tabs, 1):
                print(f"  {i}. {tab}")

        location = validate_location(self.api, filepath, read_excel_sheet)
        if not location:
            print("\nValidation failed. Returning to menu.")
            return None

        if not validate_webex_users_data(filepath, read_excel_sheet):
            print("\nValidation failed. Returning to menu.")
            return None

        if not validate_available_numbers(self.api, location, filepath, read_excel_sheet):
            print("\nValidation failed. Returning to menu.")
            return None

        validate_translation_pattern(self.api, location, filepath, read_excel_sheet, additional_tabs)
        validate_call_park_extensions(self.api, location, filepath, read_excel_sheet, additional_tabs)
        validate_and_create_schedules(self.api, location['id'], filepath)

        print("\nValidation complete.")
        return filepath, location, additional_tabs

    # ------------------------------------------------------------------
    # Lightweight bootstrap for Reset Store — only needs file + location,
    # skips number availability and other validations
    # ------------------------------------------------------------------
    def _reset_bootstrap(self):
        from libraries.aso_validation import validate_excel_file, validate_location

        if not os.path.exists('bulk'):
            print("Status: FAILED - 'bulk' folder not found")
            return None

        print("\nSearching for valid import files in bulk folder...")
        filepath = select_import_file()
        if not filepath:
            print("Status: FAILED - No valid import file selected.")
            return None

        print(f"Status: PASS - Using file: {filepath}")

        is_valid, additional_tabs = validate_excel_file(filepath)
        if not is_valid:
            print("\nValidation failed. Please fix the issues and try again.")
            return None

        location = validate_location(self.api, filepath, read_excel_sheet)
        if not location:
            print("\nCould not resolve location. Returning to menu.")
            return None

        return filepath, location

    # ------------------------------------------------------------------
    # Option 2: Reset Store
    # ------------------------------------------------------------------
    def _run_reset_store(self):
        from libraries.reset_store import reset_store

        result = self._reset_bootstrap()
        if not result:
            return
        filepath, location = result

        reset_store(self.api, location, filepath, read_excel_sheet)

    # ------------------------------------------------------------------
    # Standalone: Workspace Import only
    # ------------------------------------------------------------------
    def _run_workspace_import(self):
        result = self._aso_bootstrap()
        if not result:
            return
        filepath, location, _ = result
        process_bulk_import(self.api, location, filepath)

    # ------------------------------------------------------------------
    # Standalone: Side Car Speed Dials Import only
    # ------------------------------------------------------------------
    def _run_sidecar_import(self):
        from libraries.workspace_config import configure_side_car_speed_dials

        result = self._aso_bootstrap()
        if not result:
            return
        filepath, location, _ = result

        workspace_map, data_rows = self._build_workspace_map(filepath, location)
        if workspace_map is None:
            return

        configure_side_car_speed_dials(self.api, workspace_map, data_rows, filepath, read_excel_sheet)

    # ------------------------------------------------------------------
    # Standalone: Hunt Group Import only
    # ------------------------------------------------------------------
    def _run_huntgroup_import(self):
        from libraries.configure_hunt_groups import configure_hunt_groups

        result = self._aso_bootstrap()
        if not result:
            return
        filepath, location, _ = result

        workspace_map, data_rows = self._build_workspace_map(filepath, location)
        if workspace_map is None:
            return

        configure_hunt_groups(self.api, location, workspace_map, data_rows, filepath)

    # ------------------------------------------------------------------
    # Standalone: Auto Attendant Import only
    # ------------------------------------------------------------------
    def _run_call_handler_import(self):
        from libraries.configure_auto_attendant import configure_auto_attendants

        result = self._aso_bootstrap()
        if not result:
            return
        filepath, location, _ = result

        configure_auto_attendants(self.api, location, filepath, read_excel_sheet)

    # ------------------------------------------------------------------
    # Option 3: Scaffold Auto Attendant (no extensions/numbers)
    # ------------------------------------------------------------------
    def _run_scaffold_auto_attendant(self):
        from libraries.configure_auto_attendant import configure_auto_attendants

        result = self._reset_bootstrap()
        if not result:
            return
        filepath, location = result

        configure_auto_attendants(self.api, location, filepath, read_excel_sheet,
                                  scaffold_only=True)

    # ------------------------------------------------------------------
    # Helper: resolve extension -> workspace ID map from Control Hub
    # ------------------------------------------------------------------
    def _build_workspace_map(self, filepath, location):
        users_data = read_excel_sheet(filepath, 'Webex Users')
        if not users_data or len(users_data) < 2:
            print("Error: Could not read Webex Users sheet.")
            return None, None

        data_rows = users_data[1:]

        print(f"\nFetching existing workspaces for location '{location['name']}'...")
        ws_result = self.api.call(
            "GET", "workspaces",
            params={"orgId": self.api.org_id, "locationId": location['id'], "max": 1000}
        )
        if "error" in ws_result:
            print(f"Error fetching workspaces: {ws_result['error']}")
            return None, None

        name_to_id = {
            ws.get('displayName', '').strip(): ws.get('id')
            for ws in ws_result.get('items', [])
        }

        workspace_map = {}
        for row_idx, row in enumerate(data_rows, start=2):
            user_type = str(row[9]).strip().lower() if len(row) > 9 else ""
            if user_type == 'user':
                continue
            display_name = str(row[12]).strip() if len(row) > 12 else ""
            if display_name in name_to_id:
                workspace_map[row_idx] = name_to_id[display_name]
            else:
                print(f"  Warning: Row {row_idx} - workspace '{display_name}' not found in Control Hub, skipping.")

        if not workspace_map:
            print("No matching workspaces found. Ensure workspaces have been imported first.")
            return None, None

        print(f"Resolved {len(workspace_map)} workspace(s).")
        return workspace_map, data_rows

    # ------------------------------------------------------------------
    # Option 4: View Organization License Usage
    # ------------------------------------------------------------------
    def _run_view_license_usage(self):
        from libraries.view_license_usage import view_license_usage

        view_license_usage(self.api)

    # ------------------------------------------------------------------
    # Option 5: Reassign Workspace Calling Subscription
    # ------------------------------------------------------------------
    def _run_reassign_workspace_licenses(self):
        from libraries.reassign_workspace_licenses import reassign_workspace_licenses

        reassign_workspace_licenses(self.api)

    def main_menu(self):
        print(f"\nWelcome to Webex Control Hub CLI")
        print(f"Organization ID: {self.org_id}")
        print(f"Session ID: {self.session_id}")
        
        while True:
            choice = self.display_menu(
                "Main Menu",
                [
                    "ASO Bulk Import Tool (All in One)",
                    "Reset Store",
                    "Scaffold Auto Attendant",
                    "View Organization License Usage",
                    "Reassign Workspace Calling Subscription",
                    "Exit"
                ],
                show_back=False
            )
            
            if choice == "/b" or choice == "6":
                print("\nExiting...")
                print("Session ended")
                self.cleanup()
                break
            elif choice == "1":
                aso_bulk_import_tool(self.api)
                input("\nPress Enter to continue...")
            elif choice == "2":
                self._run_reset_store()
                input("\nPress Enter to continue...")
            elif choice == "3":
                self._run_scaffold_auto_attendant()
                input("\nPress Enter to continue...")
            elif choice == "4":
                self._run_view_license_usage()
                input("\nPress Enter to continue...")
            elif choice == "5":
                self._run_reassign_workspace_licenses()
                input("\nPress Enter to continue...")
            else:
                print("Invalid choice. Please try again.")
    
    def cleanup(self):
        try:
            self.cli_log_file.close()
            sys.stdout = sys.__stdout__
            sys.stderr = sys.__stderr__
        except:
            pass

def main():
    cli = None
    try:
        cli = WebexCLI()
        cli.main_menu()
    except KeyboardInterrupt:
        print("\n\nInterrupted by user. Exiting...")
    except Exception as e:
        print(f"\nUnexpected error: {e}")
    finally:
        if cli:
            cli.cleanup()

if __name__ == "__main__":
    main()
