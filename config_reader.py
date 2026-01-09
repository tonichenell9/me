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
        
        # Multi-line values for email_body use [END] marker
        email_body = Hi all,
        This is line 2.
        Kind regards
        [END]
    
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
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        i = 0
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            
            # Skip empty lines and comments
            if not stripped or stripped.startswith('#'):
                i += 1
                continue
            
            # Check if this is a key = value line
            if '=' in line:
                parts = line.split('=', 1)
                key = parts[0].strip()
                value = parts[1].strip() if len(parts) > 1 else ''
                
                # Special handling for email_body (multi-line)
                if key == 'email_body':
                    body_lines = [value]
                    i += 1
                    # Read until we hit [END], a comment, or another key=value
                    while i < len(lines):
                        next_line = lines[i]
                        next_stripped = next_line.strip()
                        
                        # Stop at [END] marker
                        if next_stripped == '[END]':
                            i += 1
                            break
                        # Stop at comments
                        if next_stripped.startswith('#'):
                            break
                        # Stop at new key = value (but not if line starts with space)
                        if '=' in next_line and not next_line.startswith(' ') and not next_line.startswith('\t'):
                            # Check if it looks like a key (no spaces before =)
                            potential_key = next_line.split('=')[0].strip()
                            if ' ' not in potential_key:
                                break
                        
                        body_lines.append(next_line.rstrip())
                        i += 1
                    
                    # Join and clean up the body
                    value = '\n'.join(body_lines).strip()
                    save_config_value(config, key, value)
                else:
                    save_config_value(config, key, value)
                    i += 1
            else:
                i += 1
        
        # Validate required fields
        validate_config(config)
        
        return config
        
    except Exception as e:
        print(f"Error reading configuration file: {str(e)}")
        sys.exit(1)


def save_config_value(config: dict, key: str, value):
    """Save a configuration value to the config dictionary."""
    
    # Handle boolean values
    if isinstance(value, str):
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
    
    # Distribution list (legacy - can be comma-separated or single)
    elif key == 'distribution_list':
        if isinstance(value, str):
            # Split by comma
            emails = [e.strip() for e in value.split(',')]
            config['distribution_list'] = [e for e in emails if e]
    
    # To recipients
    elif key == 'to_recipients':
        if isinstance(value, str):
            emails = [e.strip() for e in value.split(',')]
            config['to_recipients'] = [e for e in emails if e]
    
    # CC recipients
    elif key == 'cc_recipients':
        if isinstance(value, str):
            emails = [e.strip() for e in value.split(',')]
            config['cc_recipients'] = [e for e in emails if e]
    
    # Other settings
    else:
        config[key] = value


def validate_config(config: dict):
    """Validate that required configuration is present."""
    
    errors = []
    
    # Check for to_recipients or legacy distribution_list
    if not config.get('to_recipients') and not config.get('distribution_list'):
        errors.append("to_recipients is required")
    
    if not config.get('sender_name'):
        errors.append("sender_name is required")
    
    if errors:
        print("Configuration errors:")
        for error in errors:
            print(f"  - {error}")
        print("\nPlease edit config.txt and fill in the required values.")
        sys.exit(1)
    
    # If using legacy distribution_list, copy to to_recipients
    if not config.get('to_recipients') and config.get('distribution_list'):
        config['to_recipients'] = config['distribution_list']


if __name__ == "__main__":
    # Test the config reader
    config = read_config()
    print("Configuration loaded successfully!")
    print(f"Sender: {config.get('sender_name')}")
    print(f"Recipients: {config.get('distribution_list')}")
    print(f"Email body: {config.get('email_body', '')[:50]}...")
