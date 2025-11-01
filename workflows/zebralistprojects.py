"""
Zebra MicroStrategy Project Lister

This script demonstrates how to list and manage projects in the Zebra MicroStrategy environment.
It uses the connection created by ZebraCreateConnection.py and properly closes it afterwards
using ZebraCloseConnection.py.

This script shows project management capabilities and settings for administrators.
"""

import sys
import os
import csv
import json
from datetime import datetime

# Add the workflows directory to path to import other modules
sys.path.append(os.path.dirname(__file__))

from mstrio.server import Environment

# Try to import pandas for Excel export, fallback to CSV if not available
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False


def flatten_json_field(json_data, field_prefix=""):
    """
    Flatten a JSON object into individual fields.
    
    Args:
        json_data: JSON object to flatten
        field_prefix (str): Prefix for the field names
        
    Returns:
        dict: Flattened fields
    """
    flattened = {}
    
    if json_data is None or json_data == 'N/A':
        return {f"{field_prefix}_id": "N/A", f"{field_prefix}_name": "N/A", f"{field_prefix}_fullname": "N/A"}
    
    try:
        # If it's a string, try to parse as JSON
        if isinstance(json_data, str):
            try:
                json_data = json.loads(json_data)
            except json.JSONDecodeError:
                # If it fails, treat as a simple string
                return {f"{field_prefix}_name": json_data, f"{field_prefix}_id": "N/A", f"{field_prefix}_fullname": "N/A"}
        
        # If it's a dictionary, extract common fields
        if isinstance(json_data, dict):
            flattened[f"{field_prefix}_id"] = json_data.get('id', 'N/A')
            flattened[f"{field_prefix}_name"] = json_data.get('name', 'N/A')
            flattened[f"{field_prefix}_fullname"] = json_data.get('fullName', json_data.get('full_name', 'N/A'))
            flattened[f"{field_prefix}_username"] = json_data.get('username', json_data.get('userName', 'N/A'))
            flattened[f"{field_prefix}_email"] = json_data.get('email', json_data.get('emailAddress', 'N/A'))
            flattened[f"{field_prefix}_type"] = json_data.get('type', 'N/A')
            flattened[f"{field_prefix}_subtype"] = json_data.get('subtype', json_data.get('subType', 'N/A'))
        else:
            # If it's some other type, convert to string
            flattened[f"{field_prefix}_value"] = str(json_data)
            flattened[f"{field_prefix}_id"] = "N/A"
            flattened[f"{field_prefix}_name"] = "N/A"
            
    except Exception as e:
        # If all else fails, set default values
        flattened[f"{field_prefix}_id"] = "N/A"
        flattened[f"{field_prefix}_name"] = str(json_data) if json_data else "N/A"
        flattened[f"{field_prefix}_error"] = f"Parse error: {str(e)}"
    
    return flattened


def get_zebra_connection():
    """
    Get connection from ZebraCreateConnection.py or create a new one.
    
    Returns:
        Connection: MicroStrategy connection object
    """
    try:
        # Try to get existing connection from ZebraCreateConnection
        from ZebraCreateConnection import zebra_connection
        if zebra_connection:
            print("✓ Using existing connection from ZebraCreateConnection")
            return zebra_connection
    except (ImportError, AttributeError, NameError):
        print("No existing connection found, creating new connection...")
    
    # Create new connection if no existing one found
    try:
        # Import and run ZebraCreateConnection to create connection
        import ZebraCreateConnection
        conn = ZebraCreateConnection.main()
        if conn:
            print("✓ New connection created successfully")
            return conn
        else:
            raise Exception("Failed to create connection")
    except Exception as e:
        print(f"✗ Error creating connection: {e}")
        return None


def close_zebra_connection(conn):
    """
    Close the connection directly using the connection object.
    
    Args:
        conn: MicroStrategy connection object to close
    """
    try:
        if conn:
            print("\nClosing MicroStrategy connection...")
            
            # Check if connection is still active before closing
            if conn.status():
                print("✓ Connection is active, proceeding with closure...")
            else:
                print("⚠ Connection appears to be inactive, but attempting to close...")
            
            # Close the connection directly
            conn.close()
            print("✓ Connection closed successfully!")
        else:
            print("⚠ No connection object to close")
    except Exception as e:
        print(f"✗ Error closing connection: {e}")
        # Fallback: Try using ZebraCloseConnection script
        try:
            print("Attempting fallback connection cleanup...")
            import ZebraCloseConnection
            ZebraCloseConnection.main()
        except Exception as fallback_error:
            print(f"✗ Fallback cleanup also failed: {fallback_error}")


def collect_project_data(conn):
    """
    Collect project data silently and return as a list of dictionaries.
    
    Args:
        conn: MicroStrategy connection object
        
    Returns:
        list: List of project dictionaries with all relevant information
    """
    projects_data = []
    
    try:
        print("Collecting Zebra MicroStrategy Projects data...")
        
        # Create environment object
        env = Environment(connection=conn)
        
        # Get list of all projects
        all_projects = env.list_projects()
        
        if all_projects:
            total_projects = len(all_projects)
            print(f"Found {total_projects} projects. Processing silently...")
            
            # Show progress every 50 projects for large datasets
            progress_interval = max(1, total_projects // 10) if total_projects > 20 else total_projects
            
            for i, project in enumerate(all_projects, 1):
                # Show progress for large datasets
                if i % progress_interval == 0 or i == total_projects:
                    print(f"Processing: {i}/{total_projects} projects...")
                
                project_data = {
                    'index': i,
                    'name': project.name,
                    'id': project.id,
                    'description': getattr(project, 'description', 'N/A'),
                    'status': getattr(project, 'status', 'N/A'),
                    'date_created': 'N/A',
                    'date_modified': 'N/A',
                    'properties_count': 0
                }
                
                # Try to get additional details (silently)
                try:
                    project_data['date_created'] = str(getattr(project, 'date_created', 'N/A'))
                    project_data['date_modified'] = str(getattr(project, 'date_modified', 'N/A'))
                    
                    # Get owner information and flatten JSON structure
                    owner_data = getattr(project, 'owner', None)
                    owner_flattened = flatten_json_field(owner_data, "owner")
                    project_data.update(owner_flattened)
                    
                    # Get properties count (silently)
                    try:
                        properties = project.list_properties()
                        project_data['properties_count'] = len(properties) if properties else 0
                    except Exception:
                        project_data['properties_count'] = 0
                        
                except Exception:
                    # Silently continue if some details can't be retrieved
                    # Add default owner fields
                    default_owner = flatten_json_field(None, "owner")
                    project_data.update(default_owner)
                
                projects_data.append(project_data)
            
            print(f"✓ Data collection completed! Processed {len(projects_data)} projects.")
        else:
            print("No projects found in the environment.")
            
    except Exception as e:
        print(f"✗ Error collecting project data: {e}")
    
    return projects_data


def export_to_excel(projects_data, filename=None):
    """
    Export project data to Excel file using pandas.
    
    Args:
        projects_data (list): List of project dictionaries
        filename (str, optional): Custom filename for the Excel file
        
    Returns:
        str: Path to the created Excel file
    """
    if not projects_data:
        print("No project data to export.")
        return None
    
    if not filename:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"zebra_projects_{timestamp}.xlsx"
    
    try:
        # Ensure .xlsx extension
        if not filename.endswith('.xlsx'):
            filename += '.xlsx'
        
        # Get the full path
        filepath = os.path.join(os.path.dirname(__file__), filename)
        
        if PANDAS_AVAILABLE:
            # Create DataFrame
            df = pd.DataFrame(projects_data)
            
            # Rename columns for better readability
            df.rename(columns={
                'index': 'Index',
                'name': 'Project Name',
                'id': 'Project ID',
                'description': 'Description',
                'status': 'Status',
                'owner_id': 'Owner ID',
                'owner_name': 'Owner Name',
                'owner_fullname': 'Owner Full Name',
                'owner_username': 'Owner Username',
                'owner_email': 'Owner Email',
                'owner_type': 'Owner Type',
                'owner_subtype': 'Owner Subtype',
                'date_created': 'Date Created',
                'date_modified': 'Date Modified',
                'properties_count': 'Properties Count'
            }, inplace=True)
            
            # Export to Excel with formatting
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Zebra Projects', index=False)
                
                # Get the worksheet to apply formatting
                worksheet = writer.sheets['Zebra Projects']
                
                # Auto-adjust column widths
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            print(f"✓ Project data exported to Excel: {filepath}")
            return filepath
        else:
            print("⚠ pandas not available, falling back to CSV export...")
            return export_to_csv(projects_data, filename.replace('.xlsx', '.csv'))
        
    except Exception as e:
        print(f"✗ Error exporting to Excel: {e}")
        print("Falling back to CSV export...")
        return export_to_csv(projects_data, filename.replace('.xlsx', '.csv'))


def export_to_csv(projects_data, filename=None):
    """
    Export project data to CSV file.
    
    Args:
        projects_data (list): List of project dictionaries
        filename (str, optional): Custom filename for the CSV
        
    Returns:
        str: Path to the created CSV file
    """
    if not projects_data:
        print("No project data to export.")
        return None
    
    if not filename:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"zebra_projects_{timestamp}.csv"
    
    try:
        # Ensure .csv extension
        if not filename.endswith('.csv'):
            filename += '.csv'
        
        # Get the full path
        filepath = os.path.join(os.path.dirname(__file__), filename)
        
        # Define CSV headers
        headers = ['Index', 'Project Name', 'Project ID', 'Description', 'Status', 
                  'Owner ID', 'Owner Name', 'Owner Full Name', 'Owner Username', 'Owner Email', 
                  'Owner Type', 'Owner Subtype', 'Date Created', 'Date Modified', 'Properties Count']
        
        # Write to CSV
        with open(filepath, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            
            # Write header
            writer.writerow(headers)
            
            # Write data
            for project in projects_data:
                writer.writerow([
                    project['index'],
                    project['name'],
                    project['id'],
                    project['description'],
                    project['status'],
                    project.get('owner_id', 'N/A'),
                    project.get('owner_name', 'N/A'),
                    project.get('owner_fullname', 'N/A'),
                    project.get('owner_username', 'N/A'),
                    project.get('owner_email', 'N/A'),
                    project.get('owner_type', 'N/A'),
                    project.get('owner_subtype', 'N/A'),
                    project['date_created'],
                    project['date_modified'],
                    project['properties_count']
                ])
        
        print(f"✓ Project data exported to: {filepath}")
        return filepath
        
    except Exception as e:
        print(f"✗ Error exporting to CSV: {e}")
        return None


def process_and_export_projects(conn):
    """
    Process all projects silently and export directly to Excel.
    
    Args:
        conn: MicroStrategy connection object
        
    Returns:
        str: Path to the exported file
    """
    # Collect project data silently
    projects_data = collect_project_data(conn)
    
    if projects_data:
        print(f"\nFound {len(projects_data)} projects. Exporting to Excel...")
        
        # Generate filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"zebra_projects_{timestamp}.xlsx"
        
        # Export to Excel
        excel_path = export_to_excel(projects_data, filename)
        
        if excel_path:
            print(f"✓ Export completed successfully!")
            print(f"File location: {excel_path}")
            return excel_path
        else:
            print("✗ Export failed.")
            return None
    else:
        print("No project data to export.")
        return None


def main():
    """
    Main function to process Zebra MicroStrategy projects and export to Excel.
    """
    print("=== Zebra Project Exporter ===")
    
    # Get connection
    conn = get_zebra_connection()
    
    if conn:
        try:
            # Process and export projects silently
            export_path = process_and_export_projects(conn)
            
            if export_path:
                print(f"\n✓ Process completed successfully!")
                print(f"Projects exported to: {os.path.basename(export_path)}")
            else:
                print("\n✗ Export process failed.")
            
        except Exception as e:
            print(f"✗ Error during project processing: {e}")
        finally:
            # Always try to close connection
            close_zebra_connection(conn)
    else:
        print("✗ Failed to establish connection. Cannot process projects.")


if __name__ == "__main__":
    main()