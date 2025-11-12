# Zebra MicroStrategy Configuration System

This directory contains configuration management for Zebra MicroStrategy automation scripts.

## Files Overview

### Configuration Files

- **`zebra_config.json`** - Template configuration with placeholder values (safe to commit)
- **`zebra_config_local.json`** - Local configuration with actual credentials (DO NOT COMMIT)
- **`zebra_config_manager.py`** - Configuration management module
- **`__init__.py`** - Package initialization

### Configuration Priority

The configuration manager automatically selects configuration files in this order:

1. `zebra_config_local.json` (if exists) - Local configuration with real credentials
2. `zebra_config.json` (fallback) - Template configuration with placeholders

## Setup Instructions

### 1. Initial Setup

Copy the template configuration and add your credentials:

```bash
cp zebra_config.json zebra_config_local.json
```

### 2. Update Credentials

Edit `zebra_config_local.json` and replace placeholder values:

```json
{
  "microstrategy": {
    "username": "your_actual_username",
    "password": "your_actual_password"
  }
}
```

### 3. Git Safety

The `.gitignore` file is configured to prevent accidentally committing credentials:

- ✅ `zebra_config.json` (template) - SAFE to commit
- ❌ `zebra_config_local.json` (credentials) - NEVER commit
- ❌ `zebra_config_prod.json` (production) - NEVER commit

## Configuration Structure

```json
{
  "microstrategy": {
    "base_url": "https://your-mstr-server/reporting/api",
    "username": "your_username",
    "password": "your_password",
    "ssl_verify": false,
    "timeout": 300
  },
  "projects": {
    "default": "project_id_1",
    "analytics": "project_id_2"
  },
  "folders": {
    "freeform_sql": "folder_id_for_ffsql_reports"
  },
  "datasources": {
    "new_dbi_object_id": "new_datasource_object_id"
  },
  "export": {
    "output_directory": "./outputs",
    "excel_format": {
      "index": false,
      "engine": "openpyxl"
    }
  }
}
```

## Usage in Scripts

### Basic Usage

```python
from config.zebra_config_manager import ZebraConfig

# Initialize configuration
config = ZebraConfig()

# Get connection parameters
url = config.get_base_url()
username = config.get_username()
password = config.get_password()
```

### With Connection Helper

```python
import ZebraCreateConnection

# Create connection using configuration
conn = ZebraCreateConnection.create_connection(use_config=True)
```

### Direct Configuration Access

```python
# Get specific values
project_id = config.get_default_project_id()
folder_id = config.get_freeform_sql_folder_id()
dbi_id = config.get_new_dbi_object_id()

# Using dot notation
custom_value = config.get_config('projects.analytics')
```

## Validation

Test your configuration:

```python
python -m config.zebra_config_manager
```

This will:
- Load the configuration
- Validate credentials are set
- Display configuration summary (without showing sensitive values)

## Environment-Specific Configurations

For different environments, create separate configuration files:

- `zebra_config_dev.json` - Development environment
- `zebra_config_staging.json` - Staging environment  
- `zebra_config_prod.json` - Production environment

Then specify the configuration file explicitly:

```python
config = ZebraConfig('config/zebra_config_prod.json')
```

## Security Best Practices

1. **Never commit credential files** - Always use `.gitignore`
2. **Use environment variables** for CI/CD pipelines
3. **Rotate credentials regularly**
4. **Use service accounts** for automation
5. **Restrict permissions** on configuration files

## Troubleshooting

### Common Issues

1. **"Configuration file not found"**
   - Ensure `zebra_config_local.json` exists or `zebra_config.json` is present
   - Check file permissions

2. **"Please update your credentials"**
   - Replace `YOUR_USERNAME_HERE` and `YOUR_PASSWORD_HERE` in your local config
   - Verify credentials are correct

3. **"KeyError: Configuration key not found"**
   - Check JSON syntax in configuration file
   - Ensure all required keys are present
   - Compare with template structure

### Debugging

Enable debug mode to see which configuration file is being used:

```python
config = ZebraConfig()
config.print_config_summary()
```

This will show which configuration file was loaded and whether credentials are configured.

## Migration from Hardcoded Values

The configuration system automatically extracts these previously hardcoded values:

- MicroStrategy server URL
- Username and password
- Project IDs
- Folder IDs for specific operations
- Datasource/DBI object IDs

All zebra*.py scripts have been updated to use this configuration system instead of hardcoded values.