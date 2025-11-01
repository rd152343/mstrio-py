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


def create_connection(base_url, username, password, project_id=None, project_name=None):
    """
    Create a connection to MicroStrategy environment.
    
    Args:
        base_url (str): URL of the MicroStrategy REST API server
        username (str): Username for authentication
        password (str): Password for authentication
        project_id (str, optional): Project ID to connect to
        project_name (str, optional): Project name to connect to
        
    Returns:
        Connection: MicroStrategy connection object if successful, None otherwise
    """
    try:
        # Create connection object
        conn = Connection(
            base_url=base_url,
            username=username,
            password=password,
            project_id=project_id,
            project_name=project_name,
            ssl_verify=False  # Set to False for self-signed certificates or dev environments
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
    
    # Zebra MicroStrategy environment details
    ZEBRA_URL = "https://zwc-dev-usc1.cloud.microstrategy.com/reporting/api"
    ZEBRA_USERNAME = "svc_dev_usc1_admin"
    ZEBRA_PASSWORD = "P6*KBHaT%Hn5"
    ZEBRA_PROJECT_ID = "3FAB3265F7483C928678B6BF0564D92A"
    
    # Create connection
    conn = create_connection(
        base_url=ZEBRA_URL,
        username=ZEBRA_USERNAME,
        password=ZEBRA_PASSWORD,
        project_id=ZEBRA_PROJECT_ID
    )

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