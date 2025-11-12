"""
Zebra Change Project Ownership Script

This script lists all projects and allows changing project ownership in MicroStrategy.

Author: Zebra Technologies
Date: November 2025
"""

import json
import sys
import os

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


def get_zebra_connection():
    """Get connection using ZebraCreateConnection (no project needed for projects API)."""
    try:
        conn = ZebraCreateConnection.create_connection(
            use_config=True,
            project_id=None  # No project needed for projects API
        )
        if conn:
            print("✓ Connection established")
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
        success = ZebraCloseConnection.close_connection(conn)
        if success:
            print("✓ Connection closed")
    except Exception as e:
        print(f"✗ Error closing connection: {e}")


def get_all_projects(conn):
    """Get list of all projects with their details."""
    try:
        print("📋 Fetching all projects...")
        
        # API endpoint to get all projects
        endpoint = "/api/projects"
        
        response = conn.get(endpoint=endpoint)
        
        if response.status_code == 200:
            response_data = response.json()
            projects = response_data if isinstance(response_data, list) else response_data.get('projects', [])
            
            print(f"📈 Found {len(projects)} projects")
            print(f"\n{'='*80}")
            print(f"{'#':<3} {'Project Name':<30} {'Project ID':<32} {'Owner Name':<20} {'Owner ID'}")
            print(f"{'='*80}")
            
            for i, project in enumerate(projects, 1):
                project_name = project.get('name', 'Unknown')
                project_id = project.get('id', 'Unknown')
                owner = project.get('owner', {})
                owner_name = owner.get('name', 'Unknown')
                owner_id = owner.get('id', 'Unknown')
                
                print(f"{i:<3} {project_name[:29]:<30} {project_id:<32} {owner_name[:19]:<20} {owner_id}")
            
            print(f"{'='*80}")
            return projects
        else:
            print(f"❌ API Error {response.status_code}: {response.text}")
            return []
            
    except Exception as e:
        print(f"❌ Error fetching projects: {e}")
        return []


def change_project_ownership(conn, project_id, new_owner_id):
    """Change ownership of a specific project."""
    try:
        print(f"🔄 Changing ownership for project: {project_id}")
        
        # Step 1: Get current project details
        endpoint = f"/api/projects/{project_id}"
        print("📋 Getting current project details...")
        
        get_response = conn.get(endpoint=endpoint)
        if get_response.status_code != 200:
            print(f"❌ Failed to get project details: {get_response.status_code} - {get_response.text}")
            return False
        
        project_data = get_response.json()
        print("✅ Retrieved current project details")
        
        # Step 2: Update the ownerId in the complete project structure
        project_data['ownerId'] = new_owner_id
        
        # Step 3: Send PATCH with complete project structure
        headers = conn.headers.copy()
        headers['Content-Type'] = 'application/json'
        headers['Accept'] = 'application/json'
        
        print("🔄 Updating project ownership...")
        patch_response = conn.patch(endpoint=endpoint, headers=headers, json=project_data)
        
        if patch_response.status_code == 200:
            print(f"✅ Project ownership changed successfully")
            return True
        else:
            print(f"❌ Failed to change project ownership: {patch_response.status_code} - {patch_response.text}")
            return False
            
    except Exception as e:
        print(f"❌ Error changing project ownership: {e}")
        return False


def main():
    """Main function to change ownership of Projects."""
    print("=== Zebra Change Project Ownership Script ===")
    print("Lists all projects and allows changing project ownership")
    
    conn = None
    
    try:
        # Establish connection
        conn = get_zebra_connection()
        if not conn:
            print("❌ Failed to establish connection. Exiting.")
            return
        
        # Get all projects
        projects = get_all_projects(conn)
        
        if not projects:
            print("❌ No projects found")
            return
        
        # Ask user which project to change ownership for
        print(f"\nSelect project to change ownership (1-{len(projects)}):")
        try:
            choice = int(input("Enter project number: ")) - 1
            if choice < 0 or choice >= len(projects):
                print("❌ Invalid project number")
                return
        except ValueError:
            print("❌ Please enter a valid number")
            return
        
        selected_project = projects[choice]
        project_name = selected_project['name']
        project_id = selected_project['id']
        current_owner = selected_project.get('owner', {})
        current_owner_name = current_owner.get('name', 'Unknown')
        current_owner_id = current_owner.get('id', 'Unknown')
        
        print(f"\n📋 Selected Project:")
        print(f"   Name: {project_name}")
        print(f"   ID: {project_id}")
        print(f"   Current Owner: {current_owner_name} ({current_owner_id})")
        
        # Get new owner ID
        new_owner_id = input("\nEnter NEW Owner ID: ").strip()
        if not new_owner_id:
            print("❌ New Owner ID is required!")
            return
        
        # Ask for confirmation
        print(f"\n⚠️  About to change project ownership:")
        print(f"   Project: {project_name}")
        print(f"   From: {current_owner_name} ({current_owner_id})")
        print(f"   To: {new_owner_id}")
        
        confirm = input("\nProceed with ownership change? (y/n): ").lower()
        if confirm != 'y':
            print("❌ Operation cancelled")
            return
        
        # Change project ownership
        print(f"\n{'='*60}")
        print("CHANGING PROJECT OWNERSHIP")
        print(f"{'='*60}")
        
        success = change_project_ownership(conn, project_id, new_owner_id)
        
        if success:
            print(f"✅ Project ownership changed successfully!")
            print(f"   Project: {project_name}")
            print(f"   New Owner ID: {new_owner_id}")
        else:
            print(f"❌ Failed to change project ownership")
        
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