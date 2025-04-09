"""
Utility functions for the WAVE UI automation project.
"""
import os
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
        
        save_dialog = Desktop(backend="win32").window(class_name=SAVE_DIALOG_CLASS, 
                                                           title=SAVE_DIALOG_TITLE)
        
        save_dialog.wait('visible', timeout=DEFAULT_TIMEOUT)
        
        if not save_dialog.exists():
            raise RuntimeError(f"Save As dialog not found after waiting {DEFAULT_TIMEOUT} seconds")
            
        logger.info("Save As dialog found")
        
        # Find the filename field using the values from UI_FIELDS
        try:
            filename_field = save_dialog.child_window(
                auto_id=UI_FIELDS["file_name"]["auto_id"], 
                class_name=UI_FIELDS["file_name"]["class_name"]
            )
            
            if filename_field.exists():
                filename_field.set_focus()
                keyboard.send_keys("^a")
                keyboard.send_keys("{DEL}")
                keyboard.send_keys(file_path)
                logger.info(f"Entered file path: {file_path}")
            else:
                raise RuntimeError("Filename field not found by specified values")
        except Exception as e:
            logger.error(f"Error using specified values for filename field: {e}. Using keyboard fallback.")
            # Fallback to keyboard input directly
            save_dialog.set_focus()
            keyboard.send_keys("^a")
            keyboard.send_keys("{DEL}")
            keyboard.send_keys(file_path)
            logger.info(f"Entered file path: {file_path}")
        
        # Click the Save button using the values from UI_FIELDS
        try:
            save_button = save_dialog.child_window(
                auto_id=UI_FIELDS["save_button"]["auto_id"], 
                class_name=UI_FIELDS["save_button"]["class_name"]
            )
            
            if save_button.exists():
                save_button.click_input()
            else:
               raise RuntimeError("Save button is not found")
        except Exception as e:
            logger.error(f"Error clicking Save button: {e}. Using Enter key instead.")
            keyboard.send_keys("{ENTER}")
        
        # Handle overwrite confirmation if needed
        time.sleep(SAVE_FILE_TIMEOUT)
        try:
            confirm_dialog = Desktop(backend="win32").window(title_re=CONFIRM_DIALOG_TITLE)
            if confirm_dialog.exists():
                confirm_dialog.set_focus()
                keyboard.send_keys("y")  # Press 'Y' for Yes
                logger.info("Confirmed file overwrite.")
                time.sleep(SAVE_FILE_TIMEOUT)
        except Exception as e:
            logger.warning(f"No confirm dialog or error handling it: {e}")
        
        logger.info(f"File save operation completed for: {file_path}")
        return True
    
    except Exception as e:
        logger.info(f"Error handling Save As dialog: {e}")
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