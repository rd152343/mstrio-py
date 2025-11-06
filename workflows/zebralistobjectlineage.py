"""
Zebra MicroStrategy Object Analyzer

Comprehensive script to:
1. Get all Attributes, Metrics, and Facts from a project
2. Call REST API for detailed information on each object
3. Flatten the JSON responses 
4. Export to Excel with separate tabs for each object type

Author: Zebra Technologies
Date: November 2025
"""

import sys
import os

# Add the workflows directory to path to import other modules
sys.path.append(os.path.dirname(__file__))

try:
    from mstrio.object_management import list_objects
    from mstrio.types import ObjectTypes
    from mstrio.object_management.search_enums import SearchDomain
    from mstrio.connection import Connection
    import requests
    import json
    import pandas as pd
    from datetime import datetime
    print("✓ Successfully imported mstrio modules")
except ImportError as e:
    print(f"✗ Error importing mstrio modules: {e}")
    sys.exit(1)


def create_connection(project_id):
    """Create connection to MicroStrategy environment."""
    try:
        # Zebra MicroStrategy environment details
        ZEBRA_URL = "https://zwc-dev-usc1.cloud.microstrategy.com/reporting/api"
        ZEBRA_USERNAME = "svc_dev_usc1_admin"
        ZEBRA_PASSWORD = "P6*KBHaT%Hn5"
        
        # Disable SSL warnings
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        
        # Create connection
        conn = Connection(
            base_url=ZEBRA_URL,
            username=ZEBRA_USERNAME,
            password=ZEBRA_PASSWORD,
            project_id=project_id,
            ssl_verify=False
        )
        
        print("✓ Connection established successfully")
        return conn
        
    except Exception as e:
        print(f"✗ Error creating connection: {e}")
        return None


def get_object_details(conn, object_id, object_type):
    """Get detailed object information via REST API."""
    try:
        # Determine endpoint based on object type
        if object_type == 'ATTRIBUTES':
            endpoint = f"/api/model/attributes/{object_id}"
        elif object_type == 'METRICS':
            endpoint = f"/api/model/metrics/{object_id}"
        elif object_type == 'FACTS':
            endpoint = f"/api/model/facts/{object_id}"
        else:
            print(f"  ✗ Unknown object type: {object_type}")
            return None
            
        full_url = f"{conn.base_url}{endpoint}"
        print(f"  📡 API Call: {full_url}")
        
        response = conn.get(endpoint=endpoint)
        print(f"  📊 Response Status: {response.status_code}")
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"  ✗ API Error {response.status_code}: {response.text}")
            return None
            
    except Exception as e:
        print(f"  ✗ Error calling API for {object_id}: {e}")
        return None


def flatten_json(data, parent_key='', sep='_'):
    """Flatten nested JSON structure."""
    items = []
    if isinstance(data, dict):
        for k, v in data.items():
            new_key = f"{parent_key}{sep}{k}" if parent_key else k
            if isinstance(v, dict):
                items.extend(flatten_json(v, new_key, sep=sep).items())
            elif isinstance(v, list):
                for i, item in enumerate(v):
                    if isinstance(item, dict):
                        items.extend(flatten_json(item, f"{new_key}{sep}{i}", sep=sep).items())
                    else:
                        items.append((f"{new_key}{sep}{i}", item))
            else:
                items.append((new_key, v))
    elif isinstance(data, list):
        for i, item in enumerate(data):
            if isinstance(item, dict):
                items.extend(flatten_json(item, f"{parent_key}{sep}{i}", sep=sep).items())
            else:
                items.append((f"{parent_key}{sep}{i}", item))
    else:
        items.append((parent_key, data))
    
    return dict(items)


def process_objects_and_export(conn, project_id):
    """Process all object types, get detailed info via API, and export to Excel."""
    
    object_types = [
        ('ATTRIBUTES', ObjectTypes.ATTRIBUTE),
        ('METRICS', ObjectTypes.METRIC),
        ('FACTS', ObjectTypes.FACT)
    ]
    
    all_data = {}
    
    for type_name, object_type in object_types:
        print(f"\n=== Processing {type_name} ===")
        
        try:
            objects = list_objects(
                connection=conn,
                object_type=object_type,
                project_id=project_id,
                domain=SearchDomain.PROJECT
            )
            
            if objects:
                print(f"Found {len(objects)} {type_name.lower()}")
                
                # Process each object and get detailed info
                flattened_data = []
                
                for i, obj in enumerate(objects, 1):
                    print(f"  Processing {i}/{len(objects)}: {obj.id}")
                    
                    # Get detailed info via REST API
                    obj_details = get_object_details(conn, obj.id, type_name)
                    
                    if obj_details:
                        # Flatten the JSON response
                        flattened = flatten_json(obj_details)
                        # Add basic object info
                        flattened['object_id'] = obj.id
                        flattened['object_name'] = getattr(obj, 'name', 'Unknown')
                        flattened_data.append(flattened)
                    else:
                        # Add basic info even if API call failed
                        flattened_data.append({
                            'object_id': obj.id,
                            'object_name': getattr(obj, 'name', 'Unknown'),
                            'api_error': 'Failed to get details'
                        })
                
                all_data[type_name] = flattened_data
                print(f"  ✓ Processed {len(flattened_data)} {type_name.lower()}")
                
            else:
                print(f"No {type_name.lower()} found")
                all_data[type_name] = []
                
        except Exception as e:
            print(f"Error processing {type_name.lower()}: {e}")
            all_data[type_name] = []
    
    # Export to Excel
    if any(all_data.values()):
        export_to_excel(all_data, project_id)
    else:
        print("No data to export")


def export_to_excel(all_data, project_id):
    """Export flattened data to Excel with separate tabs."""
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"MicroStrategy_Objects_{project_id}_{timestamp}.xlsx"
        
        print(f"\n📊 Exporting to Excel: {filename}")
        
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            for object_type, data in all_data.items():
                if data:
                    df = pd.DataFrame(data)
                    # Fill NaN values with empty string for better readability
                    df = df.fillna('')
                    df.to_excel(writer, sheet_name=object_type, index=False)
                    print(f"  ✓ {object_type}: {len(data)} rows, {len(df.columns)} columns")
                else:
                    # Create empty sheet if no data
                    pd.DataFrame().to_excel(writer, sheet_name=object_type, index=False)
                    print(f"  ✓ {object_type}: Empty sheet created")
        
        print(f"\n🎉 Export completed: {filename}")
        
        # Print summary
        print(f"\n📈 Export Summary:")
        for object_type, data in all_data.items():
            print(f"   {object_type}: {len(data)} objects")
        
        return filename
        
    except Exception as e:
        print(f"✗ Error exporting to Excel: {e}")
        return None


def main():
    """Main function."""
    print("=== Zebra Object ID Lister ===")
    
    # Get project ID
    print("\n💡 Default Zebra project ID: 3FAB3265F7483C928678B6BF0564D92A")
    project_id = input("Enter Project ID (or press Enter for default): ").strip()
    
    if not project_id:
        project_id = "3FAB3265F7483C928678B6BF0564D92A"
        print(f"✓ Using default project: {project_id}")
    
    # Create connection
    conn = create_connection(project_id)
    if not conn:
        print("✗ Failed to establish connection. Exiting.")
        return
    
    try:
        # Process objects and export to Excel
        process_objects_and_export(conn, project_id)
        
    except Exception as e:
        print(f"✗ Error: {e}")
        
    finally:
        # Close connection
        if conn:
            conn.close()
            print("\n✓ Connection closed")


if __name__ == "__main__":
    main()
