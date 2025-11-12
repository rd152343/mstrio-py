import json
import sys
import os
from mstrio.types import ObjectTypes

# Add the workflows directory to path to import other modules
sys.path.append(os.path.dirname(__file__))

try:
    import ZebraCreateConnection
    import ZebraCloseConnection
    from config.zebra_config_manager import ZebraConfig
    print("✓ Successfully imported Zebra connection modules and configuration")
except ImportError as e:
    print(f"✗ Error importing modules: {e}")
    sys.exit(1)


def get_zebra_connection(project_id=None):
    """Get connection using ZebraCreateConnection."""
    try:
        print("Creating MicroStrategy connection...")
        
        # Load configuration for project ID if not provided
        if project_id is None:
            config = ZebraConfig()
            project_id = config.get_analytics_project_id()
        
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
    """Get all reports from a specific folder."""
    try:
        print(f"📁 Fetching reports from Folder ID: {folder_id}")
        endpoint = f"/api/folders/{folder_id}"
        response = conn.get(endpoint=endpoint)
        
        if response.status_code == 200:
            response_data = response.json()
            report_ids = []
            
            if isinstance(response_data, list):
                for item in response_data:
                    if item.get('type') == ObjectTypes.REPORT_DEFINITION.value:
                        report_ids.append({
                            'id': item.get('id'),
                            'name': item.get('name', 'Unknown')
                        })
                        print(f"   📊 Found Report: {item.get('name')} (ID: {item.get('id')})")
            
            print(f"📈 Total reports found: {len(report_ids)}")
            return report_ids
        else:
            print(f"❌ API Error {response.status_code}: {response.text}")
            return []
    except Exception as e:
        print(f"❌ Error fetching reports from folder: {e}")
        return []


def update_report_dbi(conn, report_id, report_name, base_url, new_dbi_object_id):
    """Update DBI object ID for a specific report."""
    try:
        print(f"\n🔄 Processing Report: {report_name} (ID: {report_id})")
        
        # 1. Create report instance
        instance_url = f'{base_url}/api/model/reports/{report_id}/instances'
        instance_response = conn.post(instance_url, headers=conn.headers)
        if not instance_response.ok:
            raise RuntimeError(f"Failed to create report instance: {instance_response.text}")

        instance_data = instance_response.json()
        instance_id = instance_data.get('id')
        if not instance_id:
            raise RuntimeError("Could not find instance ID in response.")

        print(f"✅ Created report instance: {instance_id}")

        # 2. Prepare headers with X-MSTR-MS-Instance
        custom_headers = conn.headers.copy()
        custom_headers['X-MSTR-MS-Instance'] = instance_id

        # 3. Get report definition
        report_def_url = f"{base_url}/api/model/reports/{report_id}?showFilterTokens=true&showExpressionAs=tree&showAdvancedProperties=true"
        def_response = conn.get(report_def_url)
        if not def_response.ok:
            raise RuntimeError(f"Failed to get report definition: {def_response.text}")

        report_json = def_response.json()
        print(f"✅ Retrieved report definition")

        # 4. Update the DBI object ID
        try:
            old_dbi_id = report_json['dataSource']['table']['dataSource']['objectId']
            report_json['dataSource']['table']['dataSource']['objectId'] = new_dbi_object_id
            print(f"🔄 Updated DBI: {old_dbi_id} → {new_dbi_object_id}")
        except KeyError:
            raise RuntimeError("Could not find DBI path. Check the report JSON structure.")

        updated_json_data = json.dumps(report_json)

        # 5. PUT updated definition
        put_url = f"{base_url}/api/model/reports/{report_id}"
        put_response = conn.put(put_url, headers=custom_headers, data=updated_json_data)
        if not put_response.ok:
            raise RuntimeError(f"DBI update failed: {put_response.text}")

        print(f"✅ Updated report definition")

        # 6. Save the report instance
        save_url = f"{base_url}/api/model/reports/{report_id}/instances/save"
        save_response = conn.post(save_url, headers=custom_headers)
        if not save_response.ok:
            raise RuntimeError(f"Instance save failed: {save_response.text}")

        print(f"✅ Saved report instance")

        # 7. (Optional) Verify the update succeeded
        final_url = f"{base_url}/api/model/reports/{report_id}"
        final_response = conn.get(final_url, headers=conn.headers)
        if not final_response.ok:
            raise RuntimeError(f"Final fetch failed: {final_response.text}")

        print(f"✅ DBI update completed for {report_name}")
        return True

    except Exception as e:
        print(f"❌ Error updating {report_name}: {e}")
        return False


def main():
    """Main function to update DBI for all reports in folder."""
    print("=== Zebra FFSQL DataSource Change Script ===")
    
    # Load configuration
    config = ZebraConfig()
    FOLDER_ID = config.get_freeform_sql_folder_id()
    NEW_DBI_OBJECT_ID = config.get_new_dbi_object_id()
    
    print(f"Target Folder ID: {FOLDER_ID}")
    print(f"New DBI Object ID: {NEW_DBI_OBJECT_ID}")
    
    conn = None
    
    try:
        # Get project ID (optional)
        project_response = input("\nEnter Project ID (press Enter for default): ").strip()
        project_id = project_response if project_response else None
        
        # Establish connection
        conn = get_zebra_connection(project_id)
        if not conn:
            print("❌ Failed to establish connection. Exiting.")
            return
        
        # Get all reports from folder
        reports = get_reports_from_folder(conn, FOLDER_ID)
        if not reports:
            print("❌ No reports found in the specified folder")
            return
        
        # Process each report
        successful_updates = 0
        failed_updates = 0
        
        print(f"\n{'='*80}")
        print("PROCESSING DBI UPDATES FOR ALL REPORTS")
        print(f"{'='*80}")
        
        for report in reports:
            report_id = report['id']
            report_name = report['name']
            
            success = update_report_dbi(conn, report_id, report_name, conn.base_url, NEW_DBI_OBJECT_ID)
            if success:
                successful_updates += 1
            else:
                failed_updates += 1
        
        # Final summary
        print(f"\n{'='*80}")
        print("FINAL SUMMARY")
        print(f"{'='*80}")
        print(f"✅ Successful Updates: {successful_updates}")
        print(f"❌ Failed Updates: {failed_updates}")
        print(f"📊 Total Reports Processed: {len(reports)}")
        
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