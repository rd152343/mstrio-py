"""
Zebra Configuration Manager

This module provides centralized configuration management for Zebra MicroStrategy scripts.
It reads configuration from zebra_config.json and provides methods to access settings.

Usage:
    from config.zebra_config_manager import ZebraConfig
    
    config = ZebraConfig()
    url = config.get_base_url()
    username = config.get_username()
    password = config.get_password()
"""

import json
import os
from typing import Dict, Any, Optional


class ZebraConfig:
    """Configuration manager for Zebra MicroStrategy scripts."""
    
    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize configuration manager.
        
        Args:
            config_path (str, optional): Path to config file. If None, tries to find suitable config.
        """
        if config_path is None:
            # Try to find a suitable config file
            script_dir = os.path.dirname(__file__)
            
            # Priority order: local -> default
            config_candidates = [
                os.path.join(script_dir, 'zebra_config_local.json'),  # Local config with actual credentials
                os.path.join(script_dir, 'zebra_config.json')         # Default template config
            ]
            
            config_path = None
            for candidate in config_candidates:
                if os.path.exists(candidate):
                    config_path = candidate
                    break
            
            if config_path is None:
                config_path = config_candidates[-1]  # Use default if none found
        
        self.config_path = config_path
        self._config: Dict[str, Any] = {}
        self._load_config()
    
    def _load_config(self) -> None:
        """Load configuration from JSON file."""
        try:
            if not os.path.exists(self.config_path):
                raise FileNotFoundError(f"Configuration file not found: {self.config_path}")
            
            with open(self.config_path, 'r', encoding='utf-8') as file:
                self._config = json.load(file)
                
            print(f"✓ Configuration loaded from: {self.config_path}")
            
        except FileNotFoundError as e:
            print(f"✗ Configuration file not found: {e}")
            print("Please ensure zebra_config.json exists in the config directory")
            raise
        except json.JSONDecodeError as e:
            print(f"✗ Invalid JSON in configuration file: {e}")
            raise
        except Exception as e:
            print(f"✗ Error loading configuration: {e}")
            raise
    
    def get_config(self, key_path: str, default: Any = None) -> Any:
        """
        Get configuration value using dot notation.
        
        Args:
            key_path (str): Dot-separated path to config value (e.g., 'microstrategy.base_url')
            default (Any): Default value if key not found
            
        Returns:
            Any: Configuration value or default
        """
        keys = key_path.split('.')
        value = self._config
        
        try:
            for key in keys:
                value = value[key]
            return value
        except (KeyError, TypeError):
            if default is not None:
                return default
            raise KeyError(f"Configuration key not found: {key_path}")
    
    # MicroStrategy connection settings
    def get_base_url(self) -> str:
        """Get MicroStrategy base URL."""
        return self.get_config('microstrategy.base_url')
    
    def get_username(self) -> str:
        """Get MicroStrategy username."""
        username = self.get_config('microstrategy.username')
        if username == "YOUR_USERNAME_HERE":
            raise ValueError(
                "Please update the username in zebra_config.json. "
                "Replace 'YOUR_USERNAME_HERE' with your actual username."
            )
        return username
    
    def get_password(self) -> str:
        """Get MicroStrategy password."""
        password = self.get_config('microstrategy.password')
        if password == "YOUR_PASSWORD_HERE":
            raise ValueError(
                "Please update the password in zebra_config.json. "
                "Replace 'YOUR_PASSWORD_HERE' with your actual password."
            )
        return password
    
    def get_ssl_verify(self) -> bool:
        """Get SSL verification setting."""
        return self.get_config('microstrategy.ssl_verify', False)
    
    def get_timeout(self) -> int:
        """Get connection timeout."""
        return self.get_config('microstrategy.timeout', 300)
    
    # Project settings
    def get_default_project_id(self) -> str:
        """Get default project ID."""
        return self.get_config('projects.default')
    
    def get_analytics_project_id(self) -> str:
        """Get analytics project ID."""
        return self.get_config('projects.analytics')
    
    def get_project_id(self, project_name: str = 'default') -> str:
        """
        Get project ID by name.
        
        Args:
            project_name (str): Project name ('default' or 'analytics')
            
        Returns:
            str: Project ID
        """
        return self.get_config(f'projects.{project_name}')
    
    # Folder settings
    def get_freeform_sql_folder_id(self) -> str:
        """Get freeform SQL folder ID."""
        return self.get_config('folders.freeform_sql')
    
    # Datasource settings  
    def get_new_dbi_object_id(self) -> str:
        """Get new DBI object ID for datasource changes."""
        return self.get_config('datasources.new_dbi_object_id')
    
    # Export settings
    def get_output_directory(self) -> str:
        """Get output directory for exports."""
        return self.get_config('export.output_directory', './outputs')
    
    def get_excel_settings(self) -> Dict[str, Any]:
        """Get Excel export settings."""
        return self.get_config('export.excel_format', {
            'index': False,
            'engine': 'openpyxl'
        })
    
    def validate_credentials(self) -> bool:
        """
        Validate that credentials have been updated from defaults.
        
        Returns:
            bool: True if credentials are valid, False otherwise
        """
        try:
            self.get_username()
            self.get_password()
            return True
        except ValueError:
            return False
    
    def print_config_summary(self) -> None:
        """Print a summary of current configuration (without sensitive data)."""
        print("\n=== Zebra Configuration Summary ===")
        print(f"Base URL: {self.get_base_url()}")
        print(f"Username: {'***CONFIGURED***' if self.get_config('microstrategy.username') != 'YOUR_USERNAME_HERE' else 'NOT_CONFIGURED'}")
        print(f"Password: {'***CONFIGURED***' if self.get_config('microstrategy.password') != 'YOUR_PASSWORD_HERE' else 'NOT_CONFIGURED'}")
        print(f"SSL Verify: {self.get_ssl_verify()}")
        print(f"Timeout: {self.get_timeout()}s")
        print(f"Default Project: {self.get_default_project_id()}")
        print(f"Analytics Project: {self.get_analytics_project_id()}")
        print(f"Output Directory: {self.get_output_directory()}")
        print("=" * 38)


# Convenience function for quick access
def get_zebra_config() -> ZebraConfig:
    """Get a ZebraConfig instance."""
    return ZebraConfig()


if __name__ == "__main__":
    # Test the configuration
    try:
        config = ZebraConfig()
        config.print_config_summary()
        
        if not config.validate_credentials():
            print("\n⚠ WARNING: Please update your credentials in zebra_config.json")
            print("Replace 'YOUR_USERNAME_HERE' and 'YOUR_PASSWORD_HERE' with actual values")
        else:
            print("\n✓ Configuration is valid and ready to use")
            
    except Exception as e:
        print(f"✗ Configuration error: {e}")