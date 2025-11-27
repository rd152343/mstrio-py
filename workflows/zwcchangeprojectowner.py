"""
Zebra Change Project Owner

This script changes the owner of a MicroStrategy project.

Author: Zebra Technologies
Date: November 2025
"""

import sys
import os

# Add the workflows directory to path to import other modules
sys.path.append(os.path.dirname(__file__))

try:
    from mstrio.server import Project
    from mstrio.users_and_groups import User
    print("✓ Successfully imported mstrio modules")
except ImportError as e:
    print(f"✗ Error importing mstrio modules: {e}")
    sys.exit(1)

# Import connection modules
try:
    from ZwcCreateConnection import create_connection
    from ZwcCloseConnection import close_connection
    print("✓ Successfully imported Zebra connection modules")
except ImportError as e:
    print(f"✗ Error importing Zebra connection modules: {e}")
    print("Make sure ZwcCreateConnection.py and ZwcCloseConnection.py are in the same directory")
    sys.exit(1)


def get_zebra_connection():
    """Get connection using ZwcCreateConnection."""
    try:
        conn = create_connection(use_config=True)
        if conn:
            print("✓ Connection established successfully")
            return conn
        else:
            print("✗ Failed to establish connection")
            return None
    except Exception as e:
        print(f"✗ Error creating connection: {e}")
        return None


def close_zebra_connection(conn):
    """Close the connection."""
    try:
        close_connection(conn)
        print("✓ Connection closed successfully")
    except Exception as e:
        print(f"✗ Error closing connection: {e}")


def change_project_owner(conn, project_id, new_owner_id):
    """
    Change the owner of a MicroStrategy project.
    
    Args:
        conn: MicroStrategy connection object
        project_id (str): ID of the project to modify
        new_owner_id (str): ID of the new owner user
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        print(f"\n=== Changing Project Owner ===")
        print(f"Project ID: {project_id}")
        print(f"New Owner ID: {new_owner_id}")
        
        # Get the project object
        print("\nFetching project...")
        project = Project(connection=conn, id=project_id)
        print(f"✓ Project found: {project.name}")
        print(f"  Current owner: {project.owner['name']} (ID: {project.owner['id']})")
        
        # Get the new owner user
        print(f"\nFetching new owner user...")
        new_owner = User(connection=conn, id=new_owner_id)
        print(f"✓ User found: {new_owner.name} (ID: {new_owner.id})")
        
        # Change the owner
        print(f"\nChanging project owner from {project.owner['name']} to {new_owner.name}...")
        project.alter(owner=new_owner)
        
        # Verify the change
        project.fetch()
        if project.owner['id'] == new_owner_id:
            print(f"✓ Project owner changed successfully!")
            print(f"  New owner: {project.owner['name']} (ID: {project.owner['id']})")
            return True
        else:
            print(f"✗ Owner change may not have been applied correctly")
            return False
            
    except Exception as e:
        print(f"✗ Error changing project owner: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Main function to change project owner."""
    print("=== Zebra Change Project Owner ===")
    
    # Default values for testing (remove after testing)
    DEFAULT_PROJECT_IDS = "4AE1EBB7CE43DB292DCCB69D71087375"  # Comma-separated project IDs
    DEFAULT_OWNER_ID = "E0C84128A242676FC0F956AA790DFEA2"  # New owner user ID
    
    # Get connection
    conn = get_zebra_connection()
    if not conn:
        print("✗ Failed to establish connection. Exiting.")
        return
    
    try:
        # Get project IDs from user (comma-separated)
        project_ids_input = input(f"\nEnter Project ID(s) (comma-separated for multiple) [Default: {DEFAULT_PROJECT_IDS}]: ").strip()
        if not project_ids_input:
            project_ids_input = DEFAULT_PROJECT_IDS
            print(f"Using default Project ID(s): {DEFAULT_PROJECT_IDS}")
        
        # Parse and clean project IDs
        project_ids = [pid.strip() for pid in project_ids_input.split(',') if pid.strip()]
        
        if not project_ids:
            print("✗ No valid Project IDs provided")
            close_zebra_connection(conn)
            return
        
        print(f"\n✓ Found {len(project_ids)} project(s) to process")
        
        # Get new owner user ID from user
        new_owner_id = input(f"\nEnter New Owner User ID [Default: {DEFAULT_OWNER_ID}]: ").strip()
        if not new_owner_id:
            new_owner_id = DEFAULT_OWNER_ID
            print(f"Using default Owner ID: {DEFAULT_OWNER_ID}")
        
        # Process each project one by one
        success_count = 0
        failed_count = 0
        results = []
        
        for idx, project_id in enumerate(project_ids, 1):
            print(f"\n{'='*60}")
            print(f"Processing Project {idx} of {len(project_ids)}")
            print(f"{'='*60}")
            
            success = change_project_owner(conn, project_id, new_owner_id)
            
            if success:
                success_count += 1
                results.append(f"✓ {project_id}")
            else:
                failed_count += 1
                results.append(f"✗ {project_id}")
        
        # Summary
        print(f"\n{'='*60}")
        print("=== Summary ===")
        print(f"{'='*60}")
        print(f"Total projects processed: {len(project_ids)}")
        print(f"Successful: {success_count}")
        print(f"Failed: {failed_count}")
        print(f"\nDetailed results:")
        for result in results:
            print(f"  {result}")
            
    except KeyboardInterrupt:
        print("\n\n✗ Operation cancelled by user")
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Close connection
        print("\nClosing connection...")
        close_zebra_connection(conn)
        print("Done.")


if __name__ == "__main__":
    main()
