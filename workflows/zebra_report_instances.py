"""
Zebra Report Instances Fetcher

This script uses ZebraCreateConnection.py and ZebraCloseConnection.py to establish
a MicroStrategy connection and fetch report instances via REST API.

API Endpoint: /api/model/reports/{REPORT_ID}/instances

Author: Zebra Technologies
Date: November 2025
"""

import sys
import os
import json

# Add the workflows directory to path to import other modules
sys.path.append(os.path.dirname(__file__))

try:
    import ZebraCreateConnection
    import ZebraCloseConnection
    from mstrio.types import ObjectTypes
    from config.zebra_config_manager import ZebraConfig
    print("✓ Successfully imported Zebra connection modules, ObjectTypes, and configuration")
except ImportError as e:
    print(f"✗ Error importing modules: {e}")
    print("Make sure ZebraCreateConnection.py, ZebraCloseConnection.py, and config/ are in the same directory")
    print("and mstrio library is properly installed")
    sys.exit(1)


def get_zebra_connection(project_id=None):
    """Get connection using ZebraCreateConnection with optional project."""
    try:
        print("Creating MicroStrategy connection...")
        
        # Load configuration for project ID if not provided
        if project_id is None:
            config = ZebraConfig()
            project_id = config.get_default_project_id()
        
        # Create connection using configuration (all parameters from config)
        conn = ZebraCreateConnection.create_connection(
            project_id=project_id,
            use_config=True
        )
        
        if conn:
            print("✓ Connection established successfully")
            return conn
        else:
            print("✗ Failed to create connection")
            return None
            
    except Exception as e:
        print(f"✗ Error creating connection: {e}")
        return None


def close_zebra_connection(conn):
    """Close connection using ZebraCloseConnection."""
    try:
        print("\nClosing MicroStrategy connection...")
        success = ZebraCloseConnection.close_connection(conn)
        if success:
            print("✓ Connection closed successfully")
        else:
            print("⚠ Connection closure completed with warnings")
    except Exception as e:
        print(f"✗ Error closing connection: {e}")


def get_reports_from_folder(conn, folder_id):
    """
    Get all reports from a specific folder using the folder API.
    
    Args:
        conn: MicroStrategy connection object
        folder_id (str): The parent folder ID to search for reports
        
    Returns:
        list: List of report IDs found in the folder, or empty list if failed
    """
    try:
        print(f"� Fetching reports from Folder ID: {folder_id}")
        
        # Construct the folder API endpoint
        endpoint = f"/api/folders/{folder_id}"
        
        # Make the REST API call using the connection's get method
        response = conn.get(endpoint=endpoint)
        
        print(f"📊 Response Status: {response.status_code}")
        
        if response.status_code == 200:
            response_data = response.json()
            print(f"✅ Successfully fetched folder contents")
            
            # Extract report IDs from the response
            report_ids = []
            
            # Response is a direct array of objects
            if isinstance(response_data, list):
                for item in response_data:
                    # Check if item is a report (type 3 = REPORT_DEFINITION)
                    if item.get('type') == ObjectTypes.REPORT_DEFINITION.value:
                        report_id = item.get('id')
                        report_name = item.get('name', 'Unknown')
                        if report_id:
                            report_ids.append({
                                'id': report_id,
                                'name': report_name
                            })
                            print(f"   📊 Found Report: {report_name} (ID: {report_id})")
            
            print(f"📈 Total reports found: {len(report_ids)}")
            return report_ids
        else:
            print(f"❌ API Error {response.status_code}: {response.text}")
            return []
            
    except Exception as e:
        print(f"❌ Error fetching reports from folder: {e}")
        return []


def fetch_report_instances(conn, report_id):
    """
    Fetch report instances for a given report ID via REST API.
    
    Args:
        conn: MicroStrategy connection object
        report_id (str): The report ID to fetch instances for
        
    Returns:
        dict: API response containing report instances, or None if failed
    """
    try:
        print(f"🔍 Fetching instances for Report ID: {report_id}")
        
        # Construct the API endpoint
        endpoint = f"/api/model/reports/{report_id}/instances"
        # Make the REST API call using the connection's post method
        response = conn.post(endpoint=endpoint)
        
        if response.status_code == 200:
            response_data = response.json()
            print(f"✅ Retrieved existing report instances")
            return response_data
        elif response.status_code == 201:
            response_data = response.json()
            print(f"✅ New report instance created successfully")
            return response_data
        else:
            print(f"❌ API Error {response.status_code}: {response.text}")
            return None
            
    except Exception as e:
        print(f"❌ Error fetching report instances: {e}")
        return None


def display_instances_summary(report_name, report_id, instances_data):
    """Display a summary of the report instances data."""
    try:
        print(f"\n📋 Report: {report_name} (ID: {report_id})")
        print(f"{'='*60}")
        
        if not instances_data:
            print("   No instances data to display")
            return 0
        
        # Check if instances exist in the response
        if 'instances' in instances_data:
            instances = instances_data['instances']
            print(f"   Total instances found: {len(instances)}")
            
            if instances:
                # Display all instances
                for i, instance in enumerate(instances, 1):
                    instance_id = instance.get('id', 'N/A')
                    instance_name = instance.get('name', 'N/A')
                    instance_status = instance.get('status', 'N/A')
                    created_date = instance.get('dateCreated', 'N/A')
                    
                    print(f"   {i}. Instance ID: {instance_id}")
                    print(f"      Name: {instance_name}")
                    print(f"      Status: {instance_status}")
                    print(f"      Created: {created_date}")
                    print()
                
                return len(instances)
        else:
            print("   No instances found")
            return 0
            
    except Exception as e:
        print(f"❌ Error displaying instances summary: {e}")
        return 0


def process_all_reports_in_folder(conn, folder_id):
    """Process all reports in a folder and fetch their instances."""
    try:
        # First, get all report IDs from the folder
        reports = get_reports_from_folder(conn, folder_id)
        
        if not reports:
            print("❌ No reports found in the specified folder")
            return
        
        total_instances = 0
        processed_reports = 0
        
        print(f"\n{'='*80}")
        print("PROCESSING REPORT INSTANCES FOR ALL REPORTS")
        print(f"{'='*80}")
        
        # Process each report
        for report in reports:
            report_id = report['id']
            report_name = report['name']
            
            print(f"\n🔄 Processing Report: {report_name}")
            
            # Fetch instances for this report
            instances_data = fetch_report_instances(conn, report_id)
            
            # Display instances for this report
            instance_count = display_instances_summary(report_name, report_id, instances_data)
            total_instances += instance_count
            processed_reports += 1
        
        # Final summary
        print(f"\n{'='*80}")
        print("FINAL SUMMARY")
        print(f"{'='*80}")
        print(f"📊 Total Reports Processed: {processed_reports}")
        print(f"📈 Total Instances Found: {total_instances}")
        
    except Exception as e:
        print(f"❌ Error processing reports in folder: {e}")


def main():
    """Main function to fetch report instances from specific folder."""
    print("=== Zebra Report Instances Fetcher ===")
    print("Fetching all reports from Folder ID: 43E06E215B429E35A777F5869C0565AE")
    
    conn = None
    
    # Hardcoded folder ID as requested
    folder_id = "43E06E215B429E35A777F5869C0565AE"
    
    try:
        # Get project ID (optional)
        project_response = input("\nEnter Project ID (press Enter for default): ").strip()
        project_id = project_response if project_response else None
        
        # Establish connection
        conn = get_zebra_connection(project_id)
        if not conn:
            print("❌ Failed to establish connection. Exiting.")
            return
        
        # Process all reports in the specified folder
        process_all_reports_in_folder(conn, folder_id)
            
    except KeyboardInterrupt:
        print("\n⚠ Operation cancelled by user")
    except Exception as e:
        print(f"❌ Error during execution: {e}")
    finally:
        # Always close connection
        if conn:
            close_zebra_connection(conn)


if __name__ == "__main__":
    main()