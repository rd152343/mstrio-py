"""
Zebra MicroStrategy Connection Closer

This workflow demonstrates how to properly close an active MicroStrategy connection.
Use this script after you're done with operations that require the connection.

This script can close connections created by ZebraCreateConnection.py or any other
MicroStrategy connection that needs proper cleanup.
"""

from mstrio.connection import Connection
from mstrio.helpers import IServerError

# Disable SSL warnings when ssl_verify=False
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def close_connection(conn):
    """
    Close an active MicroStrategy connection.
    
    Args:
        conn (Connection): MicroStrategy connection object to close
        
    Returns:
        bool: True if connection was closed successfully, False otherwise
    """
    try:
        if conn:
            print("Closing MicroStrategy connection...")
            
            # Check if connection is still active before closing
            if conn.status():
                print("✓ Connection is active, proceeding with closure...")
            else:
                print("⚠ Connection appears to be inactive, but attempting to close...")
            
            # Close the connection
            conn.close()
            print("✓ Connection closed successfully!")
            return True
        else:
            print("✗ No connection object provided")
            return False
            
    except Exception as e:
        print(f"✗ Error closing connection: {e}")
        return False


def close_connection_by_credentials(base_url, username, password):
    """
    Create a new connection and immediately close it.
    Useful for cleanup when you don't have the original connection object.
    
    Args:
        base_url (str): URL of the MicroStrategy REST API server
        username (str): Username for authentication
        password (str): Password for authentication
        
    Returns:
        bool: True if connection was created and closed successfully, False otherwise
    """
    try:
        print("Creating temporary connection for cleanup...")
        
        # Create connection object
        conn = Connection(
            base_url=base_url,
            username=username,
            password=password,
            ssl_verify=False
        )
        
        print("✓ Temporary connection established")
        
        # Close the connection
        result = close_connection(conn)
        return result
        
    except IServerError as e:
        print(f"✗ MicroStrategy Server Error: {e}")
        return False
    except Exception as e:
        print(f"✗ Failed to create connection for cleanup: {e}")
        return False


def main():
    """
    Main function for closing MicroStrategy connections.
    """
    print("=== Zebra MicroStrategy Connection Closer ===")
    
    # Option 1: Try to close using global connection (if available)
    try:
        # Import the global connection from the create script
        import sys
        import os
        
        # Add the workflows directory to path
        sys.path.append(os.path.dirname(__file__))
        
        # Try to import the connection from ZebraCreateConnection
        try:
            from ZebraCreateConnection import zebra_connection
            if zebra_connection:
                print("Found active connection from ZebraCreateConnection...")
                success = close_connection(zebra_connection)
                if success:
                    print("✓ Connection cleanup completed!")
                    return
        except (ImportError, AttributeError, NameError):
            print("No active connection found from ZebraCreateConnection")
    except Exception as e:
        print(f"Could not access global connection: {e}")
    
    # Option 2: Close connection using credentials (fallback method)
    print("\nAttempting connection cleanup using credentials...")
    
    # Zebra MicroStrategy environment details (same as ZebraCreateConnection)
    ZEBRA_URL = "https://zwc-dev-usc1.cloud.microstrategy.com/reporting/api"
    ZEBRA_USERNAME = "svc_dev_usc1_admin"
    ZEBRA_PASSWORD = "P6*KBHaT%Hn5"
    
    success = close_connection_by_credentials(
        base_url=ZEBRA_URL,
        username=ZEBRA_USERNAME,
        password=ZEBRA_PASSWORD
    )
    
    if success:
        print("✓ Connection cleanup completed using credentials!")
    else:
        print("✗ Failed to close connection. Manual cleanup may be required.")
        print("💡 Connections typically auto-expire after the timeout period.")


if __name__ == "__main__":
    main()