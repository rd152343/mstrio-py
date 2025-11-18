"""
 MicroStrategy Object Analyzer

Comprehensive script to:
1. Get all Attributes, Metrics, and Facts from a project
2. Call REST API for detailed information on each object
3. Process attribute forms with proper structure for analysis
4. Export to Excel with separate tabs for each object type

Expected Excel Output Structure for Attributes:
- OBJECT_ID: MicroStrategy object ID
- ATTRIBUTE_NAME: Name of the attribute
- CATEGORY: Form category (ID, DESC, etc.)
- FORM_NAME: Name of the form
- EXPRESSION: Expression text (column name)
- LOGICAL_TABLE: Table name where expression is found
- IS_LOOKUP: Y if table is the lookup table for this form, N otherwise
- ATTRIBUTE_LOOKUP: Main lookup table for the attribute
- Report Display: Y if form is used in report displays
- Browse Display: Y if form is used in browse displays
- DISPLAY_FORMAT: Display format (text, number, etc.)
- DATA_TYPE: Data type (integer, fixed_length_string, etc.)
- PRECISION: Data precision
- SCALE: Data scale

Special handling:
- Form groups (child forms) are expanded to show individual child form entries
- Each expression creates separate rows for each table it references
- Lookup table identification based on form's lookupTable property

Author:  Technologies
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
    from config.zwc_config_manager import ZwcConfig
    print("✓ Successfully imported mstrio modules and configuration")
except ImportError as e:
    print(f"✗ Error importing mstrio modules: {e}")
    sys.exit(1)


def create_connection(project_id):
    """Create connection to MicroStrategy environment."""
    try:
        # Load configuration
        config = ZwcConfig()
        
        # Disable SSL warnings
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        
        # Create connection using configuration
        conn = Connection(
            base_url=config.get_base_url(),
            username=config.get_username(),
            password=config.get_password(),
            project_id=project_id,
            ssl_verify=config.get_ssl_verify()
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
            # Use the project-specific attribute endpoint
            endpoint = f"/api/model/attributes/{object_id}"
            # Add showExpressionAs parameter to get better expression details
            params = {
                'showExpressionAs': 'tree',
                'showFilterTokens': 'false'
            }
        elif object_type == 'METRICS':
            endpoint = f"/api/model/metrics/{object_id}?showExpressionAs=tokens"
            params = {}
        elif object_type == 'FACTS':
            endpoint = f"/api/model/facts/{object_id}"
            params = {}
        else:
            print(f"  ✗ Unknown object type: {object_type}")
            return None
        
        response = conn.get(endpoint=endpoint, params=params)
        
        if response.status_code == 200:
            return response.json()
        else:
            print(f"  ✗ API Error {response.status_code}: {response.text[:200]}")
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


def extract_columns_from_expression(expression):
    """Extract column references from expression - now handles the simple text format."""
    column_names = []
    
    # The expression is simply a text field containing the column name
    expr_text = expression.get('text', '')
    if expr_text:
        # The text field directly contains the column name (e.g., "PROJECT_SKEY", "DOMAIN_ID", "PROJECT_TITLE")
        column_names.append(expr_text.strip())
    
    return {
        'column_names': column_names,
        'column_ids': []  # Column IDs would be in the tables section if needed
    }


def process_attribute_forms(obj_details, object_id, object_name):
    """Process attribute forms into the expected Excel structure with standardized categories."""
    forms_data = []
    
    # Get lookup table and display information
    attribute_lookup_table = obj_details.get('attributeLookupTable', {}).get('name', '')
    displays = obj_details.get('displays', {})
    report_displays = {display.get('id'): display.get('name') for display in displays.get('reportDisplays', [])}
    browse_displays = {display.get('id'): display.get('name') for display in displays.get('browseDisplays', [])}
    
    # Process attribute forms
    forms = obj_details.get('forms', [])
    if not forms:
        return [{'OBJECT_ID': object_id, 'ATTRIBUTE_NAME': object_name, 'error': 'No forms found'}]
    
    for form in forms:
        form_id = form.get('id')
        form_name = form.get('name', '')
        form_category = form.get('category', '')
        form_type = form.get('type', '')
        form_display_format = form.get('displayFormat', '')
        is_form_group = form.get('isFormGroup', False)
        
        # Get data type information
        data_type_info = form.get('dataType', {})
        data_type = data_type_info.get('type', '')
        precision = data_type_info.get('precision', '')
        scale = data_type_info.get('scale', '')
        
        # Get lookup table for this form
        form_lookup_table = form.get('lookupTable', {}).get('name', '')
        
        # Check if this form is in report/browse displays
        is_report_display = 'Y' if form_id in report_displays else 'N'
        is_browse_display = 'Y' if form_id in browse_displays else 'N'
        
        # Skip form groups - we only want the basic forms (ID, DESC, etc.)
        if is_form_group:
            continue
        
        # POST-PROCESSING: Standardize categories
        # Logic: All ID forms remain "ID", all DESC forms remain "DESC", all others become "DESC"
        standardized_category = form_category
        if form_category.upper() == 'ID':
            standardized_category = 'ID'
        elif form_category.upper() == 'DESC':
            standardized_category = 'DESC'
        else:
            # Any other category (like compound keys, etc.) becomes DESC
            standardized_category = 'DESC'
        
        # Process regular forms (non-form-groups)
        expressions = form.get('expressions', [])
        if expressions:
            for expr in expressions:
                expression_obj = expr.get('expression', {})
                expr_text = expression_obj.get('text', '')
                
                # Process each table in the expression
                tables = expr.get('tables', [])
                if tables:
                    for table in tables:
                        table_name = table.get('name', '')
                        
                        # Determine if this table is the lookup table
                        is_lookup = 'Y' if table_name == form_lookup_table else 'N'
                        
                        row = {
                            'OBJECT_ID': object_id,
                            'ATTRIBUTE_NAME': object_name,
                            'CATEGORY': standardized_category,  # Use standardized category
                            'FORM_NAME': form_name,
                            'EXPRESSION': expr_text,
                            'LOGICAL_TABLE': table_name,
                            'IS_LOOKUP': is_lookup,
                            'ATTRIBUTE_LOOKUP': attribute_lookup_table,
                            'Report Display': is_report_display,
                            'Browse Display': is_browse_display,
                            'DISPLAY_FORMAT': form_display_format,
                            'DATA_TYPE': data_type,
                            'PRECISION': precision,
                            'SCALE': scale
                        }
                        forms_data.append(row)
                else:
                    # No tables in expression
                    row = {
                        'OBJECT_ID': object_id,
                        'ATTRIBUTE_NAME': object_name,
                        'CATEGORY': standardized_category,  # Use standardized category
                        'FORM_NAME': form_name,
                        'EXPRESSION': expr_text,
                        'LOGICAL_TABLE': '',
                        'IS_LOOKUP': 'N',
                        'ATTRIBUTE_LOOKUP': attribute_lookup_table,
                        'Report Display': is_report_display,
                        'Browse Display': is_browse_display,
                        'DISPLAY_FORMAT': form_display_format,
                        'DATA_TYPE': data_type,
                        'PRECISION': precision,
                        'SCALE': scale
                    }
                    forms_data.append(row)
        else:
            # No expressions in form
            row = {
                'OBJECT_ID': object_id,
                'ATTRIBUTE_NAME': object_name,
                'CATEGORY': standardized_category,  # Use standardized category
                'FORM_NAME': form_name,
                'EXPRESSION': '',
                'LOGICAL_TABLE': '',
                'IS_LOOKUP': 'N',
                'ATTRIBUTE_LOOKUP': attribute_lookup_table,
                'Report Display': is_report_display,
                'Browse Display': is_browse_display,
                'DISPLAY_FORMAT': form_display_format,
                'DATA_TYPE': data_type,
                'PRECISION': precision,
                'SCALE': scale
            }
            forms_data.append(row)
    
    return forms_data


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
                    # Show progress every 50 objects
                    if i % 50 == 0:
                        print(f"  Progress: {i}/{len(objects)} objects processed")
                    
                    # Get detailed info via REST API
                    obj_details = get_object_details(conn, obj.id, type_name)
                    
                    if obj_details:
                        if type_name == 'ATTRIBUTES':
                            # Special processing for attributes to create rows for each form
                            forms_rows = process_attribute_forms(obj_details, obj.id, getattr(obj, 'name', 'Unknown'))
                            flattened_data.extend(forms_rows)
                        else:
                            # Regular flattening for METRICS and FACTS
                            flattened = flatten_json(obj_details)
                            # Add basic object info
                            flattened['OBJECT_ID'] = obj.id
                            flattened['OBJECT_NAME'] = getattr(obj, 'name', 'Unknown')
                            flattened_data.append(flattened)
                    else:
                        # Add basic info even if API call failed  
                        if type_name == 'ATTRIBUTES':
                            flattened_data.append({
                                'OBJECT_ID': obj.id,
                                'ATTRIBUTE_NAME': getattr(obj, 'name', 'Unknown'),
                                'api_error': 'Failed to get details'
                            })
                        else:
                            flattened_data.append({
                                'OBJECT_ID': obj.id,
                                'OBJECT_NAME': getattr(obj, 'name', 'Unknown'),
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
    print("===  Object ID Lister ===")
    
    # Get project ID
    print("\n💡 Default  project ID: 3FAB3265F7483C928678B6BF0564D92A")
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
