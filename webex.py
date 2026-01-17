#!/usr/bin/env python3
import os
import sys
import logging
from datetime import datetime
from typing import List

from libraries.api_client import WebexAPI
from libraries.list_workspaces import list_workspaces
from libraries.view_workspace import view_workspace_details
from libraries.create_workspace import create_workspace
from libraries.update_workspace import update_workspace
from libraries.delete_workspace import delete_workspace
from libraries.bulk_create_workspaces import bulk_create_workspaces
from libraries.aso_bulk_import import aso_bulk_import_tool

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
            # Create temporary API client for org lookup
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
    
    def workspace_menu(self):
        while True:
            choice = self.display_menu(
                "Workspace Management",
                [
                    "List Workspaces",
                    "View Workspace Details",
                    "Create Workspace",
                    "Update Workspace",
                    "Delete Workspace",
                    "Bulk Create Workspaces"
                ]
            )
            
            if choice == "/b":
                break
            elif choice == "1":
                list_workspaces(self.api)
                input("\nPress Enter to continue...")
            elif choice == "2":
                view_workspace_details(self.api)
                input("\nPress Enter to continue...")
            elif choice == "3":
                create_workspace(self.api)
                input("\nPress Enter to continue...")
            elif choice == "4":
                update_workspace(self.api)
                input("\nPress Enter to continue...")
            elif choice == "5":
                delete_workspace(self.api)
                input("\nPress Enter to continue...")
            elif choice == "6":
                bulk_create_workspaces(self.api)
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
                    "ASO Bulk Import Tool",
                    "Exit"
                ]
            )
            
            if choice == "/b" or choice == "3":
                print("\nExiting...")
                print("Session ended")
                self.cleanup()
                break
            elif choice == "1":
                self.workspace_menu()
            elif choice == "2":
                aso_bulk_import_tool(self.api)
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
