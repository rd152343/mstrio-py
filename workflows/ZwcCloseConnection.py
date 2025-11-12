"""
Zebra MicroStrategy Connection Closer

This workflow demonstrates how to properly close an active MicroStrategy connection.
Use this script after you're done with operations that require the connection.

This script can close connections created by ZwcCreateConnection.py or any other
MicroStrategy connection that needs proper cleanup.
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
    from config.zwc_config_manager import ZwcConfig
    print("✓ Configuration manager imported successfully")
except ImportError as e:
    print(f"✗ Error importing configuration manager: {e}")
    print("Please ensure config/zwc_config_manager.py exists")


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


def close_connection_by_credentials(base_url=None, username=None, password=None, use_config=True):
    """
    Create a new connection and immediately close it.
    Useful for cleanup when you don't have the original connection object.
    
    Args:
        base_url (str, optional): URL of the MicroStrategy REST API server. If None, uses config.
        username (str, optional): Username for authentication. If None, uses config.
        password (str, optional): Password for authentication. If None, uses config.
        use_config (bool): Whether to use configuration file for missing parameters
        
    Returns:
        bool: True if connection was created and closed successfully, False otherwise
    """
    try:
        print("Creating temporary connection for cleanup...")
        
        # Load configuration if needed
        if use_config and (base_url is None or username is None or password is None):
            try:
                config = ZwcConfig()
                
                # Use config values for missing parameters
                if base_url is None:
                    base_url = config.get_base_url()
                if username is None:
                    username = config.get_username()
                if password is None:
                    password = config.get_password()
                
                ssl_verify = config.get_ssl_verify()
                
            except Exception as e:
                print(f"✗ Error loading configuration: {e}")
                return False
        else:
            ssl_verify = False
        
        # Create connection object
        conn = Connection(
            base_url=base_url,
            username=username,
            password=password,
            ssl_verify=ssl_verify
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
        
        # Try to import the connection from ZwcCreateConnection
        try:
            from ZwcCreateConnection import zebra_connection
            if zebra_connection:
                print("Found active connection from ZwcCreateConnection...")
                success = close_connection(zebra_connection)
                if success:
                    print("✓ Connection cleanup completed!")
                    return
        except (ImportError, AttributeError, NameError):
            print("No active connection found from ZwcCreateConnection")
    except Exception as e:
        print(f"Could not access global connection: {e}")
    
    # Option 2: Close connection using credentials from config (fallback method)
    print("\nAttempting connection cleanup using configuration...")
    
    success = close_connection_by_credentials(use_config=True)
    
    if success:
        print("✓ Connection cleanup completed using credentials!")
    else:
        print("✗ Failed to close connection. Manual cleanup may be required.")
        print("💡 Connections typically auto-expire after the timeout period.")


if __name__ == "__main__":
    main()
