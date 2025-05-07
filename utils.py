"""
Utility functions for the WAVE UI automation project.
"""
import os
import json
import logging
import time
from pywinauto import Desktop, keyboard
from constants import (SAVE_DIALOG_CLASS, SAVE_DIALOG_TITLE, CONFIRM_DIALOG_TITLE, 
                       DEFAULT_TIMEOUT, UI_FIELDS, SAVE_FILE_TIMEOUT)

logger = logging.getLogger(__name__)

def check_and_create_results_directory(directory_path):
    """
    Ensures that the specified directory exists, creating it if necessary.
    """
    if not os.path.exists(directory_path):
        os.makedirs(directory_path, exist_ok=True)
        logger.info(f"Created directory: {directory_path}")
    return directory_path

def load_json_config(config_file):
    """
    Load and validate JSON configuration file.
    
    Args:
        config_file (str): Path to the JSON configuration file
        
    Returns:
        dict: Validated configuration dictionary
        
    Raises:
        ValueError: If the configuration is invalid
    """
    try:
        with open(config_file, 'r') as f:
            config = json.load(f)
            
        # Validate required sections
        if 'common' not in config:
            raise ValueError("Missing 'common' section in configuration")
        if 'stages' not in config:
            raise ValueError("Missing 'stages' section in configuration")
            
        # Validate common parameters
        required_common = ['wave_exe', 'project_path', 'project_name', 'case_name', 
                         'feed_flow_rate']
        for param in required_common:
            if param not in config['common']:
                raise ValueError(f"Missing required common parameter: {param}")
                
        # Validate stages
        if not config['stages'] or not isinstance(config['stages'], list):
            raise ValueError("'stages' must be a non-empty list")
            
        for i, stage in enumerate(config['stages'], 1):
            required_stage = ['element_type', 'pv_range', 'els_range']
            for param in required_stage:
                if param not in stage:
                    raise ValueError(f"Missing required parameter '{param}' in stage {i}")
                    
            # Validate ranges
            if not isinstance(stage['pv_range'], list) or len(stage['pv_range']) != 2:
                raise ValueError(f"Invalid pv_range in stage {i}: must be [min, max]")
            if not isinstance(stage['els_range'], list) or len(stage['els_range']) != 2:
                raise ValueError(f"Invalid els_range in stage {i}: must be [min, max]")
                
            # Stage-specific validation
            if i == 1:
                if 'feed_pressure_range' not in stage:
                    raise ValueError("Stage 1 must include feed_pressure_range")
                if not isinstance(stage['feed_pressure_range'], list) or len(stage['feed_pressure_range']) != 2:
                    raise ValueError("Invalid feed_pressure_range in stage 1: must be [min, max]")
            else:
                if 'target_pressure_range' not in stage:
                    raise ValueError(f"Stage {i} must include target_pressure_range")
                if not isinstance(stage['target_pressure_range'], list) or len(stage['target_pressure_range']) != 2:
                    raise ValueError(f"Invalid target_pressure_range in stage {i}: must be [min, max]")
        
        logger.info("Configuration file validated successfully")
        return config
        
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON format in config file: {e}")
    except Exception as e:
        raise ValueError(f"Error loading configuration: {e}")

def handle_save_dialog(file_path):
    """
    Handles the Save As dialog for exporting files.
    
    Args:
        file_path (str): The full path where the file should be saved.
        
    Returns:
        bool: True if successful, False otherwise.
    """
    try:
        logger.info(f"Waiting for Save As dialog to save: {file_path}")
        
        # Wait for and find the Save As dialog
        try:
            save_dialog = Desktop(backend="win32").window(class_name=SAVE_DIALOG_CLASS, 
                                                         title=SAVE_DIALOG_TITLE)
            save_dialog.wait('visible', timeout=DEFAULT_TIMEOUT)
            
            if not save_dialog.exists():
                raise RuntimeError(f"Save As dialog not found after waiting {DEFAULT_TIMEOUT} seconds")
                
            logger.info("Save As dialog found")
        except Exception as e:
            logger.error(f"Failed to find Save As dialog: {e}")
            return False
        
        # Enter the file path
        try:
            filename_field = save_dialog.child_window(
                auto_id=UI_FIELDS["file_name"]["auto_id"], 
                class_name=UI_FIELDS["file_name"]["class_name"]
            )
            
            if filename_field.exists():
                filename_field.set_focus()
                keyboard.send_keys("^a")  # Select all
                keyboard.send_keys("{DEL}")  # Delete existing text
                keyboard.send_keys(file_path, with_spaces=True)
                logger.info(f"Entered file path: {file_path}")
            else:
                logger.warning("Filename field not found by specified values, using keyboard fallback")
                save_dialog.set_focus()
                keyboard.send_keys("^a{DEL}")  # Select all and delete in one command
                keyboard.send_keys(file_path, with_spaces=True)
                logger.info(f"Entered file path using keyboard fallback: {file_path}")
        except Exception as e:
            logger.error(f"Error entering file path: {e}")
            return False
        
        # Click the Save button
        try:
            save_button = save_dialog.child_window(
                auto_id=UI_FIELDS["save_button"]["auto_id"], 
                class_name=UI_FIELDS["save_button"]["class_name"]
            )
            
            if save_button.exists():
                save_button.click_input()
                logger.info("Clicked Save button")
            else:
                logger.warning("Save button not found, using Enter key instead")
                keyboard.send_keys("{ENTER}")
                logger.info("Pressed Enter key to save")
        except Exception as e:
            logger.error(f"Error clicking Save button: {e}")
            keyboard.send_keys("{ENTER}")
            logger.info("Used Enter key as fallback")
        
        # Handle confirmation dialog if it appears
        time.sleep(SAVE_FILE_TIMEOUT)
        try:
            # Try both known dialog titles and classes for better reliability
            for title in [CONFIRM_DIALOG_TITLE, "Confirm Save As"]:
                for class_name in [SAVE_DIALOG_CLASS, "#32770"]:
                    try:
                        confirm_dialog = Desktop(backend="win32").window(title=title, class_name=class_name)
                        if confirm_dialog.exists():
                            confirm_dialog.set_focus()
                            logger.info(f"Found confirmation dialog with title '{title}' and class '{class_name}'")
                            
                            # Try multiple approaches to click Yes
                            yes_button = confirm_dialog.child_window(title="Yes", auto_id="CommandButton_6", class_name="CCPushButton")
                            if yes_button.exists():
                                yes_button.click_input()
                                logger.info("Clicked 'Yes' button in confirm dialog")
                            else:
                                # Try alternative button locators
                                yes_button = confirm_dialog.child_window(title="Yes")
                                if yes_button.exists():
                                    yes_button.click_input()
                                    logger.info("Clicked 'Yes' button by title only")
                                else:
                                    # Use keyboard shortcut as last resort
                                    logger.info("Using Alt+Y shortcut for confirm dialog")
                                    keyboard.send_keys("%y")
                            
                            # Wait for dialog to close
                            time.sleep(SAVE_FILE_TIMEOUT)
                            return True
                    except Exception as e:
                        logger.debug(f"Exception checking for confirm dialog ({title}, {class_name}): {e}")
                        continue
            
            # If we get here, no confirmation dialog was found
            logger.info("No confirmation dialog detected, assuming file saved successfully")
            return True
            
        except Exception as e:
            logger.warning(f"Error handling confirmation dialog: {e}")
            # Try pressing Alt+Y as a last resort
            keyboard.send_keys("%y")
            time.sleep(SAVE_FILE_TIMEOUT)
        
        logger.info(f"File save operation completed for: {file_path}")
        return True
    
    except Exception as e:
        logger.error(f"Error handling Save As dialog: {e}")
        return False

def find_button_by_name(window, button_name):
    """
    Find a button by its name with multiple fallback approaches.
    
    Args:
        window: The window object to search in
        button_name (str): Name of the button to find
        
    Returns:
        The button object if found, None otherwise.
    """
    try:
        # First attempt: Direct button by title
        return window.child_window(title=button_name, control_type="Button")
    except Exception:
        try:
            # Second attempt: Using child_window with class_name
            return window.child_window(title=button_name, class_name="Button")
        except Exception:
            # Third attempt: Find button by text
            buttons = window.children(class_name="Button")
            for button in buttons:
                if button_name.lower() in button.window_text().lower():
                    return button
            return None