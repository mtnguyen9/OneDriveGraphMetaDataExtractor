#!/usr/bin/env python3
"""
SharePoint File & Folder Extractor
Simple tool to list all files and folders in a SharePoint/OneDrive site with metadata
"""

import argparse
import csv
import re
import urllib.parse
from datetime import datetime
from typing import Dict, List, Optional
import requests
from msal import ConfidentialClientApplication, PublicClientApplication


class SharePointExtractor:
    def __init__(self, debug_mode: bool = False):
        self.debug_mode = debug_mode
        self.access_token = None
        self.graph_endpoint = "https://graph.microsoft.com/v1.0"
        self.total_items = 0
        self.files_count = 0
        self.folders_count = 0
        
    def log(self, message: str, level: str = "INFO"):
        """Simple logging with timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefix = {
            "INFO": "ℹ️",
            "SUCCESS": "✅",
            "WARNING": "⚠️",
            "ERROR": "❌",
            "DEBUG": "🔍"
        }
        print(f"[{timestamp}] {prefix.get(level, 'ℹ️')} {message}")
        
    def debug_log(self, message: str):
        """Debug logging - only shows if debug mode is enabled"""
        if self.debug_mode:
            self.log(message, "DEBUG")
        
    def authenticate(self, tenant_id: str = None, client_id: str = None, client_secret: str = None) -> bool:
        """Authenticate with Microsoft Graph API"""
        self.log("Starting authentication with Microsoft Graph...")
        
        # Default public client ID if none provided
        if not client_id:
            client_id = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
        
        scopes = ["https://graph.microsoft.com/Sites.Read.All", 
                 "https://graph.microsoft.com/Files.Read.All"]
        
        try:
            if client_secret and tenant_id:
                # Service principal authentication
                self.log("Using service principal authentication...")
                authority = f"https://login.microsoftonline.com/{tenant_id}"
                app = ConfidentialClientApplication(
                    client_id=client_id,
                    client_credential=client_secret,
                    authority=authority
                )
                result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
            else:
                # Interactive device flow authentication
                self.log("Using interactive device flow authentication...")
                authority = "https://login.microsoftonline.com/common"
                app = PublicClientApplication(
                    client_id=client_id,
                    authority=authority
                )
                
                # Try silent authentication first
                accounts = app.get_accounts()
                if accounts:
                    self.debug_log("Found existing account, trying silent authentication...")
                    result = app.acquire_token_silent(scopes, account=accounts[0])
                else:
                    result = None
                    
                if not result:
                    # Start device flow
                    self.log("Starting device code flow...")
                    flow = app.initiate_device_flow(scopes=scopes)
                    if "user_code" not in flow:
                        raise ValueError("Failed to create device flow")
                    
                    print(f"\n🔗 Please visit: {flow['verification_uri']}")
                    print(f"🔑 And enter this code: {flow['user_code']}\n")
                    self.log("Waiting for user authentication...")
                    result = app.acquire_token_by_device_flow(flow)
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                self.log("Authentication successful!", "SUCCESS")
                return True
            else:
                error_msg = result.get('error_description', 'Unknown authentication error')
                self.log(f"Authentication failed: {error_msg}", "ERROR")
                return False
                
        except Exception as e:
            self.log(f"Authentication error: {str(e)}", "ERROR")
            return False
    
    def parse_sharepoint_url(self, url: str) -> Dict[str, str]:
        """Parse SharePoint URL to extract site and path components"""
        self.log(f"Parsing SharePoint URL: {url}")
        
        url_info = {
            'tenant_domain': '',
            'site_path': '',
            'document_path': '',
            'is_personal_site': False
        }
        
        if not url:
            self.log("Empty URL provided", "ERROR")
            return url_info
        
        # Clean URL
        url = url.split('?')[0].rstrip('/')
        
        # Extract tenant domain
        domain_match = re.search(r'https://([^/]+)', url)
        if domain_match:
            url_info['tenant_domain'] = domain_match.group(1)
            self.debug_log(f"Extracted tenant domain: {url_info['tenant_domain']}")
        
        # Check if it's a personal OneDrive site
        personal_match = re.search(r'/personal/([^/]+)', url)
        if personal_match:
            url_info['is_personal_site'] = True
            url_info['site_path'] = f"/personal/{personal_match.group(1)}"
            self.log("Detected personal OneDrive site")
        else:
            # Team SharePoint site
            site_match = re.search(r'/sites/([^/]+)', url)
            if site_match:
                url_info['site_path'] = f"/sites/{site_match.group(1)}"
                self.log("Detected team SharePoint site")
        
        # Extract document path
        doc_match = re.search(r'/Documents/(.+)', url)
        if doc_match:
            url_info['document_path'] = urllib.parse.unquote(doc_match.group(1))
            self.log(f"Target folder: {url_info['document_path']}")
        elif '/Documents' in url:
            url_info['document_path'] = ""
            self.log("Target: Root Documents folder")
        else:
            # Handle paths without explicit Documents folder
            if url_info['is_personal_site']:
                path_match = re.search(r'/personal/[^/]+/(.+)', url)
                if path_match:
                    url_info['document_path'] = urllib.parse.unquote(path_match.group(1))
                    self.log(f"Target folder: {url_info['document_path']}")
        
        return url_info
    
    def make_graph_request(self, endpoint: str) -> Optional[Dict]:
        """Make authenticated request to Microsoft Graph API"""
        if not self.access_token:
            raise ValueError("Not authenticated")
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        
        try:
            self.debug_log(f"Making Graph API request: {endpoint}")
            response = requests.get(f"{self.graph_endpoint}{endpoint}", headers=headers)
            
            if response.status_code == 200:
                return response.json()
            elif response.status_code == 404:
                self.debug_log(f"Resource not found: {endpoint}")
                return None
            else:
                self.log(f"API request failed: {response.status_code} - {response.text}", "WARNING")
                return None
                
        except Exception as e:
            self.log(f"Request error: {str(e)}", "ERROR")
            return None
    
    def find_site(self, tenant_domain: str, site_path: str) -> Optional[str]:
        """Find and return the site ID"""
        self.log("Looking up SharePoint site...")
        
        # Try different site ID formats
        attempts = [
            f"{tenant_domain}:{site_path}",
            f"{tenant_domain}",
            "root"
        ]
        
        # Add alternative formats for SharePoint Online
        if '.' in tenant_domain:
            tenant_prefix = tenant_domain.split('.')[0]
            if "/sites/" in site_path:
                site_name = site_path.split("/sites/")[1]
                attempts.append(f"{tenant_prefix}.sharepoint.com,{site_path.replace('/', ',')[1:]}")
            elif "/personal/" in site_path:
                user_path = site_path.split("/personal/")[1]
                attempts.append(f"{tenant_prefix}.sharepoint.com,personal,{user_path}")
        
        for attempt in attempts:
            self.debug_log(f"Trying site ID: {attempt}")
            result = self.make_graph_request(f"/sites/{attempt}")
            if result:
                site_name = result.get('displayName', 'Unknown')
                self.log(f"Found site: {site_name}", "SUCCESS")
                return result['id']
        
        self.log("Could not find the specified site", "ERROR")
        return None
    
    def get_document_drive(self, site_id: str) -> Optional[str]:
        """Get the main document library drive ID"""
        self.log("Finding document library...")
        
        result = self.make_graph_request(f"/sites/{site_id}/drives")
        if not result or 'value' not in result:
            self.log("No drives found in site", "ERROR")
            return None
        
        drives = result['value']
        self.debug_log(f"Found {len(drives)} drives")
        
        # Look for Documents library first
        for drive in drives:
            if drive.get('name') == 'Documents':
                self.log("Using Documents library", "SUCCESS")
                return drive['id']
        
        # Fall back to first drive
        if drives:
            drive_name = drives[0].get('name', 'Unknown')
            self.log(f"Using drive: {drive_name}", "SUCCESS")
            return drives[0]['id']
        
        self.log("No suitable drive found", "ERROR")
        return None
    
    def find_target_folder(self, drive_id: str, path: str) -> Optional[str]:
        """Find the target folder item ID"""
        if not path:
            self.log("Starting from root folder")
            return "root"
        
        self.log(f"Navigating to folder: {path}")
        
        # Try direct path access first
        path_attempts = [
            f"root:/{path}:",
            f"root:/Documents/{path}:",
        ]
        
        for path_attempt in path_attempts:
            self.debug_log(f"Trying path: {path_attempt}")
            result = self.make_graph_request(f"/drives/{drive_id}/items/{path_attempt}")
            if result:
                self.log("Found target folder", "SUCCESS")
                return result['id']
        
        # If direct access fails, navigate step by step
        self.log("Direct path failed, navigating step by step...")
        return self.navigate_path_segments(drive_id, "root", path)
    
    def navigate_path_segments(self, drive_id: str, current_id: str, path: str) -> Optional[str]:
        """Navigate through path segments one by one"""
        segments = [seg for seg in path.split('/') if seg.strip()]
        
        for i, segment in enumerate(segments):
            self.log(f"Looking for folder: {segment} ({i+1}/{len(segments)})")
            
            # Get children of current folder
            result = self.make_graph_request(f"/drives/{drive_id}/items/{current_id}/children")
            if not result or 'value' not in result:
                self.log(f"Cannot access folder contents", "ERROR")
                return None
            
            # Find matching folder
            found = False
            for child in result['value']:
                if child['name'].lower() == segment.lower() and 'folder' in child:
                    current_id = child['id']
                    self.debug_log(f"Found: {child['name']}")
                    found = True
                    break
            
            if not found:
                available = [child['name'] for child in result['value'] if 'folder' in child]
                self.log(f"Folder '{segment}' not found. Available folders: {available}", "ERROR")
                return None
        
        self.log("Successfully navigated to target folder", "SUCCESS")
        return current_id
    
    def scan_folder_contents(self, drive_id: str, folder_id: str, current_path: str = "") -> List[Dict]:
        """Recursively scan folder contents and return file/folder list"""
        items = []
        
        try:
            # Get all items in current folder
            result = self.make_graph_request(f"/drives/{drive_id}/items/{folder_id}/children")
            if not result or 'value' not in result:
                self.debug_log(f"No items found in folder: {current_path}")
                return items
            
            children = result['value']
            self.log(f"Scanning folder: {current_path or 'Root'} ({len(children)} items)")
            
            for child in children:
                # Determine item type
                is_folder = 'folder' in child
                item_type = "Folder" if is_folder else "File"
                
                # Build full path
                item_path = f"{current_path}/{child['name']}" if current_path else child['name']
                
                # Extract metadata
                item_info = {
                    'Name': child['name'],
                    'Type': item_type,
                    'Path': item_path,
                    'Created': child.get('createdDateTime', ''),
                    'Modified': child.get('lastModifiedDateTime', ''),
                    'CreatedBy': child.get('createdBy', {}).get('user', {}).get('displayName', 'Unknown'),
                    'ModifiedBy': child.get('lastModifiedBy', {}).get('user', {}).get('displayName', 'Unknown'),
                    'Size_KB': round(child.get('size', 0) / 1024, 2) if child.get('size') else 0,
                    'IsFolder': is_folder
                }
                
                items.append(item_info)
                self.total_items += 1
                
                if is_folder:
                    self.folders_count += 1
                    self.debug_log(f"Found folder: {item_path}")
                    # Recursively scan subfolders
                    sub_items = self.scan_folder_contents(drive_id, child['id'], item_path)
                    items.extend(sub_items)
                else:
                    self.files_count += 1
                    self.debug_log(f"Found file: {item_path}")
            
        except Exception as e:
            self.log(f"Error scanning folder {current_path}: {str(e)}", "ERROR")
        
        return items
    
    def export_to_csv(self, items: List[Dict], output_file: str):
        """Export items list to CSV file"""
        if not items:
            self.log("No items to export", "WARNING")
            return
        
        self.log(f"Exporting {len(items)} items to CSV...")
        
        try:
            with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
                fieldnames = ['Name', 'Type', 'Path', 'Created', 'Modified', 
                            'CreatedBy', 'ModifiedBy', 'Size_KB', 'IsFolder']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(items)
            
            self.log(f"Export completed: {output_file}", "SUCCESS")
            
        except Exception as e:
            self.log(f"Export failed: {str(e)}", "ERROR")
    
    def run(self, sharepoint_url: str, output_file: str = None, 
            tenant_id: str = None, client_id: str = None, client_secret: str = None) -> List[Dict]:
        """Main execution method"""
        
        # Generate output filename if not provided
        if not output_file:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = f"sharepoint_contents_{timestamp}.csv"
        
        self.log("=== SharePoint File & Folder Extractor ===")
        self.log(f"Target URL: {sharepoint_url}")
        self.log(f"Output file: {output_file}")
        
        try:
            # Step 1: Authenticate
            if not self.authenticate(tenant_id, client_id, client_secret):
                raise Exception("Authentication failed")
            
            # Step 2: Parse URL
            url_info = self.parse_sharepoint_url(sharepoint_url)
            if not url_info['tenant_domain']:
                raise Exception("Could not parse SharePoint URL")
            
            # Step 3: Find site
            site_id = self.find_site(url_info['tenant_domain'], url_info['site_path'])
            if not site_id:
                raise Exception("Could not find SharePoint site")
            
            # Step 4: Get document drive
            drive_id = self.get_document_drive(site_id)
            if not drive_id:
                raise Exception("Could not access document library")
            
            # Step 5: Find target folder
            folder_id = self.find_target_folder(drive_id, url_info['document_path'])
            if not folder_id:
                raise Exception("Could not find target folder")
            
            # Step 6: Scan contents
            self.log("Starting content scan...")
            items = self.scan_folder_contents(drive_id, folder_id, url_info['document_path'])
            
            # Step 7: Export results
            self.export_to_csv(items, output_file)
            
            # Summary
            self.log("=== SCAN COMPLETE ===", "SUCCESS")
            self.log(f"Total items found: {self.total_items}")
            self.log(f"Files: {self.files_count}")
            self.log(f"Folders: {self.folders_count}")
            self.log(f"Results saved to: {output_file}")
            
            return items
            
        except Exception as e:
            self.log(f"Extraction failed: {str(e)}", "ERROR")
            return []


def main():
    parser = argparse.ArgumentParser(
        description='Extract files and folders from SharePoint/OneDrive',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python extractor.py "https://contoso.sharepoint.com/sites/myteam"
  python extractor.py "https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/Projects" -o results.csv
  python extractor.py "https://contoso.sharepoint.com/sites/myteam/Documents/Archive" --debug
        """
    )
    
    parser.add_argument('url', help='SharePoint/OneDrive URL to scan')
    parser.add_argument('-o', '--output', help='Output CSV file path')
    parser.add_argument('-t', '--tenant-id', help='Azure AD Tenant ID (for service principal)')
    parser.add_argument('-c', '--client-id', help='Azure AD Client ID')
    parser.add_argument('-s', '--client-secret', help='Azure AD Client Secret (for service principal)')
    parser.add_argument('-d', '--debug', action='store_true', help='Enable detailed debug logging')
    
    args = parser.parse_args()
    
    try:
        extractor = SharePointExtractor(debug_mode=args.debug)
        extractor.run(
            sharepoint_url=args.url,
            output_file=args.output,
            tenant_id=args.tenant_id,
            client_id=args.client_id,
            client_secret=args.client_secret
        )
        return 0
    except KeyboardInterrupt:
        print("\n⚠️  Operation cancelled by user")
        return 1
    except Exception as e:
        print(f"❌ Fatal error: {e}")
        return 1


if __name__ == "__main__":
    exit(main())