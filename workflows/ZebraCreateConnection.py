"""
Zebra MicroStrategy Connection Creator

This workflow demonstrates how to create a connection to a MicroStrategy 
environment using URL, username, and password for Zebra Technologies.

This script creates a connection and keeps it active for further operations.
Use ZebraCloseConnection.py to properly close the connection when done.
"""

from mstrio.connection import Connection
from mstrio.helpers import IServerError

# Disable SSL warnings when ssl_verify=False
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Import configuration manager
import sys
import os
sys.path.append(os.path.dirname(__file__))

try:
    from config.zebra_config_manager import ZebraConfig
    print("✓ Configuration manager imported successfully")
except ImportError as e:
    print(f"✗ Error importing configuration manager: {e}")
    print("Please ensure config/zebra_config_manager.py exists")
    sys.exit(1)


def create_connection(base_url=None, username=None, password=None, project_id=None, project_name=None, use_config=True):
    """
    Create a connection to MicroStrategy environment.
    
    Args:
        base_url (str, optional): URL of the MicroStrategy REST API server. If None, uses config.
        username (str, optional): Username for authentication. If None, uses config.
        password (str, optional): Password for authentication. If None, uses config.
        project_id (str, optional): Project ID to connect to. If None, uses config default.
        project_name (str, optional): Project name to connect to
        use_config (bool): Whether to use configuration file for missing parameters
        
    Returns:
        Connection: MicroStrategy connection object if successful, None otherwise
    """
    try:
        # Load configuration if needed
        if use_config and (base_url is None or username is None or password is None or project_id is None):
            try:
                config = ZebraConfig()
                
                # Use config values for missing parameters
                if base_url is None:
                    base_url = config.get_base_url()
                if username is None:
                    username = config.get_username()
                if password is None:
                    password = config.get_password()
                if project_id is None and project_name is None:
                    project_id = config.get_default_project_id()
                
                ssl_verify = config.get_ssl_verify()
                
                print("✓ Using configuration values for connection")
                
            except Exception as e:
                print(f"✗ Error loading configuration: {e}")
                print("Please check your zebra_config.json file")
                return None
        else:
            ssl_verify = False  # Default for manual parameters
        
        # Create connection object
        conn = Connection(
            base_url=base_url,
            username=username,
            password=password,
            project_id=project_id,
            project_name=project_name,
            ssl_verify=ssl_verify
        )
        
        print("✓ Connection established successfully!")
        
        if conn.project_id:
            print(f"Selected project: {conn.project_name} ({conn.project_id})")
        else:
            print("No project selected")
            
        return conn
        
    except IServerError as e:
        print(f"✗ MicroStrategy Server Error: {e}")
        return None
    except Exception as e:
        print(f"✗ Connection failed: {e}")
        return None


def test_connection(conn):
    """
    Test the connection by checking its status.
    
    Args:
        conn (Connection): MicroStrategy connection object
        
    Returns:
        bool: True if connection is active, False otherwise
    """
    try:
        if conn and conn.status():
            print("✓ Connection is active and healthy")
            return True
        else:
            print("✗ Connection is not active")
            return False
    except Exception as e:
        print(f"✗ Error testing connection: {e}")
        return False


def main():
    """
    Main function demonstrating connection creation.
    """
    print("=== Zebra MicroStrategy Connection Creator ===")
    
    # Create connection using configuration file
    conn = create_connection(use_config=True)

    if conn:
        # Test the connection
        test_connection(conn)
        
        # Display connection properties
        print(f"\nConnection Properties:")
        print(f"- Base URL: {conn.base_url}")
        print(f"- Username: {conn.username}")
        print(f"- User ID: {conn.user_id}")
        print(f"- Session timeout: {conn.timeout} seconds")
        
        # List available projects (if no project was selected)
        if not conn.project_id:
            try:
                from mstrio.server import list_projects
                projects = list_projects(conn, limit=5)
                print(f"\nAvailable projects (first 5):")
                for project in projects:
                    print(f"- {project.name} (ID: {project.id})")
            except Exception as e:
                print(f"Could not list projects: {e}")
        
        print("\n✓ Connection is ready for use!")
        print("💡 Use ZebraCloseConnection.py to close this connection when done.")
        
        # Return connection object for potential reuse
        global zebra_connection
        zebra_connection = conn
        return conn
    else:
        print("Failed to establish connection")
        return None


if __name__ == "__main__":
    main()