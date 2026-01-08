#!/usr/bin/env python3
"""
Simple configuration file reader.
Reads config.txt in a simple key = value format.
"""

import os
import sys


def read_config(config_path: str = 'config.txt') -> dict:
    """
    Read configuration from a simple text file.
    
    Format:
        # Comments start with #
        key = value
        
        # Multi-line values end when a new key is found
        email_body = Hi all,
        
        This is line 2.
        
        Kind regards
    
    Args:
        config_path: Path to the configuration file
        
    Returns:
        Dictionary with configuration values
    """
    
    # Check if file exists
    if not os.path.exists(config_path):
        print(f"Error: Configuration file '{config_path}' not found.")
        print("Please create a config.txt file in the same folder as this script.")
        sys.exit(1)
    
    config = {
        'email': {},
        'distribution_list': [],
    }
    
    current_key = None
    current_value_lines = []
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        for line in lines:
            # Skip empty lines at the start
            stripped = line.strip()
            
            # Skip comments
            if stripped.startswith('#'):
                continue
            
            # Check if this is a new key = value line
            if '=' in line and not line.startswith(' ') and not line.startswith('\t'):
                # Save previous key if exists
                if current_key:
                    save_config_value(config, current_key, current_value_lines)
                
                # Parse new key = value
                parts = line.split('=', 1)
                current_key = parts[0].strip()
                current_value_lines = [parts[1].strip()] if len(parts) > 1 else []
            
            elif current_key and stripped:
                # This is a continuation of the previous value (multi-line)
                current_value_lines.append(stripped)
        
        # Save the last key
        if current_key:
            save_config_value(config, current_key, current_value_lines)
        
        # Validate required fields
        validate_config(config)
        
        return config
        
    except Exception as e:
        print(f"Error reading configuration file: {str(e)}")
        sys.exit(1)


def save_config_value(config: dict, key: str, value_lines: list):
    """Save a configuration value to the config dictionary."""
    
    # Join multi-line values with newlines
    value = '\n'.join(value_lines).strip()
    
    # Handle boolean values
    if value.lower() in ('yes', 'true', '1'):
        value = True
    elif value.lower() in ('no', 'false', '0'):
        value = False
    
    # Email settings
    email_keys = [
        'incoming_subject',
        'incoming_folder',
        'previous_report_subject', 
        'previous_report_folder',
        'preview_before_send'
    ]
    
    if key in email_keys:
        config['email'][key] = value
    
    # Distribution list (can be comma-separated or single)
    elif key == 'distribution_list':
        if isinstance(value, str):
            # Split by comma or newline
            emails = [e.strip() for e in value.replace('\n', ',').split(',')]
            config['distribution_list'] = [e for e in emails if e]  # Remove empty
    
    # Other settings
    else:
        config[key] = value


def validate_config(config: dict):
    """Validate that required configuration is present."""
    
    errors = []
    
    if not config.get('distribution_list'):
        errors.append("distribution_list is required")
    
    if not config.get('sender_name'):
        errors.append("sender_name is required")
    
    if errors:
        print("Configuration errors:")
        for error in errors:
            print(f"  - {error}")
        print("\nPlease edit config.txt and fill in the required values.")
        sys.exit(1)


if __name__ == "__main__":
    # Test the config reader
    config = read_config()
    print("Configuration loaded successfully!")
    print(f"Sender: {config.get('sender_name')}")
    print(f"Recipients: {config.get('distribution_list')}")
    print(f"Previous report folder: {config['email'].get('previous_report_folder')}")
