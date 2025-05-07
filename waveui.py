"""
This class is responsible for WAVE UI automation actions
"""
from pywinauto import Application, Desktop, keyboard
import os
import time
import itertools
import logging
import pandas as pd
import json
import re
from datetime import datetime

from constants import (
    WINDOW_TITLE_PATTERN, TAB_REVERSE_OSMOSIS, TAB_SUMMARY_REPORT, 
    UI_FIELDS, BUTTON_DETAILED_REPORT, BUTTON_EXPORT, 
    TOOLBAR_ID, EXPORT_DIR,
    EXIT_DIALOG_TITLE, APPLICATION_LOAD_TIMEOUT, 
    TAB_SWITCH_TIMEOUT, DEFAULT_TIMEOUT, WINDOW_VISIBLE_TIMEOUT, SAVE_DIALOG_CLASS
)
from utils import check_and_create_results_directory, handle_save_dialog

logger = logging.getLogger(__name__)

class WaveUI:
    """
    Class for automating interactions with the DuPont WAVE software.
    """
    
    def __init__(self, file_name, project_path, project_name, case_name, feed_flow_rate=2.1, stages=1, 
                 prev_stage_excel_file=None, export_dir=EXPORT_DIR):
        """
        Initialize the WaveUI automation class.
        
        Args:
            file_name (str): Path to the WAVE executable
            project_path (str): Path to the project file
            project_name (str): Name of the project to work with
            case_name (str): Name of the case to work with
            feed_flow_rate (float): Feed flow rate value
            stages (int): Number of stages
            prev_stage_excel_file (str): Path to the previous stage excel file (used only when stages > 1)
        """
        self.file_name = file_name
        self.project_path = project_path
        self.project_name = project_name
        self.case_name = case_name
        self.feed_flow_rate = feed_flow_rate
        self.stages = stages
        self.prev_stage_excel_file = prev_stage_excel_file

        self.app = None
        self.main_window = None
        
        self.export_dir = check_and_create_results_directory(export_dir)
        self.report_counter = 1
        
        # Setup metadata tracking
        self.metadata_file = os.path.join(self.export_dir, 'report_metadata.csv')
        
        # Initialize or load existing metadata
        if os.path.exists(self.metadata_file):
            try:
                self.metadata_df = pd.read_csv(self.metadata_file)
                # Get highest report number to continue the sequence
                if not self.metadata_df.empty and 'report_id' in self.metadata_df.columns:
                    max_id = self.metadata_df['report_id'].max()
                    self.report_counter = int(max_id) + 1
                    logger.info(f"Continuing from report #{self.report_counter}")
            except Exception as e:
                logger.error(f"Error loading existing metadata file: {e}")
                self.metadata_df = self._create_empty_metadata_df()
        else:
            self.metadata_df = pd.DataFrame(columns=[
            'filename', 
            'timestamp',
            'current_stage_params', 
            'previous_stages_params'
        ])
            logger.info("Created new metadata tracking file")

    def generate_report_filename(self):
        """Generate an incremental filename for reports"""
        return f"wave_report_{self.report_counter:04d}.xls"

    def add_metadata_entry(self, stage_params, prev_stage_params=None):
        """
        Add a metadata entry for the current report.
        
        Args:
            stage_params (dict): Parameters for the current stage
            prev_stage_params (list, optional): List of parameter tuples for previous stages
            
        Returns:
            str: Generated filename
        """
        filename = self.generate_report_filename()
        
        # Convert parameters to JSON strings for easy storage and parsing
        current_params_json = json.dumps(stage_params)
        
        # Create structured data for previous stages if available
        if prev_stage_params:
            prev_stages_data = []
            for i, params in enumerate(prev_stage_params):
                # params typically is (pv, els, element_type, pressure)
                stage_num = i + 1
                prev_stage_dict = {
                    'stage': stage_num,
                    'pv': params[0],
                    'els': params[1],
                    'element_type': params[2]
                }
                
                # Add appropriate pressure fields based on stage number
                if stage_num == 1:
                    prev_stage_dict['feed_pressure'] = params[3]
                else:
                    # For stage > 1, we have boost pressure, and need to calculate target pressure
                    prev_stage_dict['target_pressure'] = params[3]
                    # If target pressure was provided in 5th position, use it
                    if len(params) > 4:
                        prev_stage_dict['boost_pressure'] = params[4]
                
                prev_stages_data.append(prev_stage_dict)
            prev_stages_json = json.dumps(prev_stages_data)
        else:
            prev_stages_json = json.dumps([])
        
        # Add entry to metadata DataFrame
        new_row = {
            'filename': filename,
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'current_stage_params': current_params_json,
            'previous_stages_params': prev_stages_json
        }
        
        self.metadata_df = pd.concat([self.metadata_df, pd.DataFrame([new_row])], ignore_index=True)
        
        # Save metadata after each addition to prevent data loss
        self.save_metadata()
        
        # Increment counter for next report
        self.report_counter += 1
        
        return filename

    def save_metadata(self):
        """Save metadata to CSV file"""
        try:
            self.metadata_df.to_csv(self.metadata_file, index=False)
            logger.debug(f"Updated metadata file: {self.metadata_file}")
            return True
        except Exception as e:
            logger.error(f"Error saving metadata: {e}")
            return False

    def launch_wave(self):
        """
        Launches the WAVE application with the specified file.
        
        Returns:
            bool: True if successful, False otherwise.
        """
        try:
            logger.info(f"Starting WAVE application with file: {self.file_name}")
            self.app = Application(backend="uia").start(f'"{self.file_name}" "{self.project_path}"')
            time.sleep(APPLICATION_LOAD_TIMEOUT)
            
            # Get the main window
            window_title = WINDOW_TITLE_PATTERN.format(
                project_name=self.project_name, 
                case_name=self.case_name
            )
            self.main_window = self.app.window(title=window_title)
            self.main_window.wait('visible', timeout=WINDOW_VISIBLE_TIMEOUT)
            logger.info("Connected to main window successfully")
            return True
        except Exception as e:
            logger.error(f"Error launching WAVE: {e}")
            return False
    
    def select_tab(self, tab_name, timeout=TAB_SWITCH_TIMEOUT):
        """
        Selects a tab in the application.
        
        Args:
            tab_name (str): The name of the tab to select
            
        Returns:
            bool: True if successful, False otherwise.
        """
        try:
            tab = self.main_window.child_window(title=tab_name, control_type="TabItem")
            tab.click_input()
            logger.info(f"Selected tab: {tab_name}")
            tab.wait('enabled', timeout=timeout)
            return True
        except Exception as e:
            logger.error(f"Failed to select tab '{tab_name}': {e}")
            return False
    
    def set_reverse_osmosis_parameters(self, pv_per_stage, els_per_pv, element_type, pressure, prev_stage_params):
        """
        Sets parameters in the Reverse Osmosis tab.
        
        Args:
            pv_per_stage (int): Number of pressure vessels per stage
            els_per_pv (int): Number of elements per pressure vessel
            pressure (int): Feed/boost pressure value based on the stage
            element_type (str): Element type of current stage
            prev_stage_params (list, None): List of tuples (pv, els, pressure) for each previous stage. Required for stages > 1.
            
        Returns:
            bool: True if successful, False otherwise.
        """
        try:
            logger.info(f"Setting RO parameters: PV={pv_per_stage}, ELS={els_per_pv}, Pressure={pressure} for stage {self.stages}")

            # Handle Stages (radio buttons)
            try:
                stages_group = self.main_window.child_window(
                    title=UI_FIELDS["stages"]["group_name"],
                    control_type="Group",
                    class_name=UI_FIELDS["stages"]["class_name"]
                )
                # Find the radio button that matches our stages value
                radio_button = stages_group.child_window(
                    title=f"{self.stages}  ",
                    control_type="RadioButton"
                )
                radio_button.click_input()
                logger.info(f"Set stages to {self.stages}")
            except Exception as e:
                logger.error(f"Failed to set stages: {e}")
        
            # Handle the fields of stages
            if self.stages == 1:
                self._set_stage_reverse_osmosis_parameters(self.stages, pv_per_stage, els_per_pv, element_type, pressure)
            else:
                logger.info(f"Previous stage(s) values are {prev_stage_params} of format (pv, els, element_type pressure, *boost_pressure)")
                cur_params = (pv_per_stage, els_per_pv, element_type, pressure)
                params = prev_stage_params + (cur_params,)
                for cur_stage, values in enumerate(params, start=1):
                    if len(values) == 4:
                        pv, els, ele_type, pressure = values
                    elif len(values) == 5:
                        pv, els, ele_type, pressure, boost_pressure = values
                        pressure = boost_pressure
                    self._set_stage_reverse_osmosis_parameters(cur_stage, pv, els, ele_type, pressure)
            return True
        except Exception as e:
            logger.error(f"Error setting RO parameters: {e}")
            return False
    
    def _set_stage_reverse_osmosis_parameters(self, cur_stage, pv_per_stage, els_per_pv, ele_type, pressure):
        
        # Handle PV per stage (edit field)
        try:
            pv_edit = self.main_window.child_window(
                auto_id=UI_FIELDS["pv_per_stage"]["auto_id"], 
                control_type="Edit",
                class_name=UI_FIELDS["pv_per_stage"]["class_name"],
                found_index=cur_stage-1
            )
            pv_edit.set_text(str(pv_per_stage))
            logger.info(f"Set PV per stage to {pv_per_stage} on stage {cur_stage}")
        except Exception as e:
            logger.error(f"Failed to set PV per stage on stage {cur_stage}: {e}")
        
        # Handle Elements per PV (edit field)
        try:
            els_edit = self.main_window.child_window(
                auto_id=UI_FIELDS["els_per_pv"]["auto_id"], 
                control_type="Edit",
                class_name=UI_FIELDS["els_per_pv"]["class_name"],
                found_index=cur_stage-1
            )
            els_edit.set_text(str(els_per_pv))
            logger.info(f"Set elements per PV to {els_per_pv} on stage {cur_stage}")
        except Exception as e:
            logger.error(f"Failed to set elements per PV on stage {cur_stage}: {e}")
        
         # Handle Element Type (combo box)
        try:
            combo_box = self.main_window.child_window(
                auto_id=UI_FIELDS["element_type"]["auto_id"][cur_stage-1],
                control_type="ComboBox",
                class_name=UI_FIELDS["element_type"]["class_name"],
            )
            combo_box.select(ele_type)
            logger.info(f"Set element type to {ele_type} on stage {cur_stage}")
        except Exception as e:
            logger.error(f"Failed to set element type on stage {cur_stage}: {e}")

        # Handle Pressure (edit field)
        if cur_stage == 1:
            # Handle the Feed pressure, only for the Stage 1
            try:
                # Use the specific AutomationId to find the feed pressure edit field
                feed_pressure_edit = self.main_window.child_window(
                    auto_id=UI_FIELDS["feed_pressure"]["auto_id"], 
                    control_type="Edit",
                    class_name=UI_FIELDS["feed_pressure"]["class_name"],
                    found_index=cur_stage-1
                )
                
                feed_pressure_edit.set_text(str(pressure))
                logger.info(f"Set feed pressure to {pressure}")
            except Exception as e:
                logger.error(f"Failed to set feed pressure: {e}")
        else:
             # Handle the boost pressure, when stage > 1
            try:
                boost_pressure_edit = self.main_window.child_window(
                    auto_id=UI_FIELDS["boost_pressure"]["auto_id"], 
                    control_type="Edit",
                    class_name=UI_FIELDS["boost_pressure"]["class_name"],
                    found_index=cur_stage-1
                )
                
                boost_pressure_edit.set_text(str(pressure))
                logger.info(f"Set boost pressure to {pressure} on stage {cur_stage}")
            except Exception as e:
                logger.error(f"Failed to set boost pressure on stage {cur_stage}: {e}")


    def open_detailed_report(self):
        """
        Opens the detailed report from the Summary Report tab.
        
        Returns:
            bool: True if successful, False otherwise.
        """
        try:
            detailed_report_button = self.main_window.child_window(
                title=BUTTON_DETAILED_REPORT, 
                control_type="Button",
                class_name=UI_FIELDS["detailed_report"]["class_name"]
            )                
            detailed_report_button.click_input()
            logger.info("Opened detailed report")
            return True
        except Exception as e:
            logger.error(f"Failed to open detailed report: {e}")
            return False
    
    def export_to_excel(self, file_path):
        """
        Exports the current report to Excel.
        
        Args:
            file_path(str): Path for the excel file
            
        Returns:
            bool: True if successful, False otherwise.
        """
        try:
            logger.info("Starting Excel export process...")
            
            # Find and click the Export button in the toolbar
            toolbar = self.main_window.child_window(auto_id=TOOLBAR_ID, control_type="ToolBar")
            export_button = toolbar.child_window(title=BUTTON_EXPORT)
            export_button.wait('ready', 15)
            export_button.click_input()
            logger.info("Clicked Export button")
             
            dropdown_item = self.main_window.child_window(title="Excel", control_type="MenuItem")
            dropdown_item.wait('ready', timeout=15)
            # Select Excel format from dropdown
            keyboard.send_keys("{DOWN}")
            keyboard.send_keys("{ENTER}")
            logger.info("Selected Excel format")
            
            
            return handle_save_dialog(file_path)
            
        except Exception as e:
            logger.error(f"Error exporting to Excel: {e}")
            return False
    
    def close_application(self):
        """
        Closes the WAVE application.
        
        Returns:
            bool: True if successful, False otherwise.
        """
        try:
            if self.main_window:
                self.main_window.close()
                self._handle_exit_confirmation()
                logger.info("Application closed successfully")
                return True
            return False
        except Exception as e:
            logger.error(f"Error closing application: {e}")
            return False
    
    def _handle_exit_confirmation(self):
        """
        Handles the exit confirmation dialog and potential save project dialog.
        
        Returns:
            bool: True if successful, False otherwise.
        """
        try:
            # Give the dialog more time to appear
            time.sleep(3)
            
            # Try to find the confirmation dialog
            try:
                confirm_dialog = Desktop(backend="win32").window(title="Confirmation", class_name="#32770")
                confirm_dialog.wait('exists', DEFAULT_TIMEOUT)
                logger.info("Found exit confirmation dialog")
                
                # Make sure dialog is in focus
                confirm_dialog.set_focus()
                time.sleep(1)
                
                # Try several approaches to click Yes
                try:
                    yes_button = confirm_dialog.child_window(title="Yes", class_name="Button")
                    yes_button.set_focus()
                    yes_button.click()
                    logger.info("Clicked Yes on exit confirmation dialog")
                except Exception as e:
                    logger.warning(f"Failed with button click: {e}")
                    try:
                        # Direct keyboard shortcut for "Yes"
                        keyboard.send_keys("%y")  # Alt+Y for Yes
                        logger.info("Sent Alt+Y for Yes button")
                        time.sleep(1)
                    except Exception as ke:
                        logger.warning(f"Failed with Alt+Y: {ke}")
                        # Last resort
                        keyboard.send_keys("{ENTER}")
                        logger.info("Pressed Enter as fallback")
            except Exception as e:
                logger.warning(f"Win32 backend failed: {e}")
                # Try UIA backend
                try:
                    confirm_dialog = Desktop(backend="uia").window(title="Confirmation")
                    confirm_dialog.wait('exists', DEFAULT_TIMEOUT)
                    logger.info("Found dialog with UIA backend")
                    
                    confirm_dialog.set_focus()
                    time.sleep(1)
                    
                    # Try to click Yes
                    yes_button = confirm_dialog.child_window(title="Yes")
                    yes_button.click()
                    logger.info("Clicked Yes with UIA backend")
                except Exception as e2:
                    logger.warning(f"UIA backend also failed: {e2}")
                    # Last resort
                    keyboard.send_keys("%y")  # Alt+Y
                    time.sleep(1)
                    keyboard.send_keys("{ENTER}")  # Enter as final fallback
                    logger.info("Used keyboard fallbacks")
            
            # Wait for dialogs to process
            time.sleep(3)
            
            # Check for save dialog
            try:
                save_dialog = Desktop(backend="win32").window(title_re="Save.*", class_name="#32770")
                if save_dialog.exists():
                    logger.info("Found save dialog")
                    save_dialog.set_focus()
                    time.sleep(1)
                    
                    try:
                        no_button = save_dialog.child_window(title="No", class_name="Button")
                        no_button.set_focus()
                        no_button.click()
                        logger.info("Clicked No on save dialog")
                    except Exception as e:
                        logger.warning(f"Failed to click No: {e}")
                        # Try keyboard shortcuts
                        keyboard.send_keys("%n")  # Alt+N for No
                        logger.info("Sent Alt+N for No button")
            except Exception as e:
                logger.debug(f"No save dialog or couldn't interact: {e}")
                # Check with UIA backend as well
                try:
                    save_dialog = Desktop(backend="uia").window(title_re="Save.*")
                    if save_dialog.exists():
                        save_dialog.set_focus()
                        time.sleep(1)
                        no_button = save_dialog.child_window(title="No")
                        no_button.click()
                        logger.info("Clicked No with UIA backend")
                except Exception as e2:
                    logger.debug(f"Save dialog check with UIA failed: {e2}")
            
            # Wait for everything to complete
            time.sleep(3)
            
            # Verify the app is actually closed
            try:
                if self.main_window.exists():
                    logger.warning("Main window still exists, sending forceful Alt+F4")
                    self.main_window.set_focus()
                    keyboard.send_keys("%{F4}")
                    time.sleep(2)
                    # Try sending escape to close any lingering dialogs
                    keyboard.send_keys("{ESC}")
                    time.sleep(1)
                    keyboard.send_keys("%y")  # Alt+Y for potential confirmation
            except Exception as e:
                logger.info(f"Main window verification error: {e}, assuming closed successfully")
            
            return True
                
        except Exception as e:
            logger.error(f"Error handling exit confirmation: {e}")
            # Emergency fallback - try brute force closing
            try:
                keyboard.send_keys("%{F4}")  # Alt+F4
                time.sleep(1)
                keyboard.send_keys("%y")  # Alt+Y for Yes
                time.sleep(1)
                keyboard.send_keys("%n")  # Alt+N for No (save dialog)
                return True
            except:
                return False

    def get_pressure_from_excel(self, stage, pv, els, element_type, pressure, prev_stage_params=None):
        """
        Get feed pressure from a specific stage's Excel report by using metadata.
        
        Args:
            stage (int): Current stage number
            pv (int): PV value for current stage
            els (int): ELS value for current stage
            element_type (str): Element type for current stage
            pressure (float): Pressure value for current stage (feed_pressure for stage 1, boost_pressure for stages > 1)
            prev_stage_params (list, optional): List of previous stage parameter tuples
            
        Returns:
            tuple: (boost_pressure, feed_pressure) both rounded to 1 decimal place if found, (float|None, float|None) otherwise
        """
        try:
            logger.info(f"Looking up feed pressure from Excel file: {self.prev_stage_excel_file}")
            
            # Read the metadata sheet from the Excel file
            metadata_df = pd.read_excel(self.prev_stage_excel_file, sheet_name='Metadata')
            logger.info(f"Loaded metadata with {len(metadata_df)} rows")
            
            # Create current stage parameters for matching
            current_params = {
                'stage': stage,
                'pv': pv, 
                'els': els,
                'element_type': element_type
            }
            
            # Add appropriate pressure field based on stage number
            if stage == 1:
                current_params['feed_pressure'] = pressure
            else:
                current_params['target_pressure'] = pressure
            
            # Create previous stage parameters list if available
            if prev_stage_params:
                prev_stages_data = []
                for i, params in enumerate(prev_stage_params):
                    # params typically is (pv, els, element_type, pressure)
                    stage_num = i + 1
                    prev_stage_dict = {
                        'stage': stage_num,
                        'pv': params[0],
                        'els': params[1],
                        'element_type': params[2]
                    }
                    
                    # Add appropriate pressure fields based on stage number
                    if stage_num == 1:
                        prev_stage_dict['feed_pressure'] = params[3]
                    else:
                        prev_stage_dict['target_pressure'] = params[3]
                        # If boost pressure was provided in 5th position, use it
                        if len(params) > 4:
                            prev_stage_dict['boost_pressure'] = params[4]
                    
                    prev_stages_data.append(prev_stage_dict)
            
            # Parse JSON strings in the DataFrame for comparison
            metadata_df['current_stage_params'] = metadata_df['current_stage_params'].apply(
                lambda x: json.loads(x) if isinstance(x, str) else x
            )
            metadata_df['previous_stages_params'] = metadata_df['previous_stages_params'].apply(
                lambda x: json.loads(x) if isinstance(x, str) else x
            )
            
            # Find matching row based on parameters
            matched_row = None
            for idx, row in metadata_df.iterrows():
                cur_params = row['current_stage_params']
                prev_params = row['previous_stages_params']
                
                # Check if basic parameters match
                if (cur_params.get('stage') == current_params['stage'] and
                    cur_params.get('pv') == current_params['pv'] and
                    cur_params.get('els') == current_params['els'] and
                    cur_params.get('element_type') == current_params['element_type']):
                    
                    # Check appropriate pressure field based on stage
                    if stage == 1:
                        if cur_params.get('feed_pressure') != current_params['feed_pressure']:
                            continue
                    else:
                        if cur_params.get('target_pressure') != current_params['target_pressure']:
                            continue
                    
                    # If previous parameters are also required, check them too
                    if prev_stage_params:
                        prev_match = True
                        if len(prev_params) != len(prev_stages_data):
                            prev_match = False
                        else:
                            for i, prev_param in enumerate(prev_params):
                                stage_num = i + 1
                                expected_param = prev_stages_data[i]
                                
                                # Check basic parameters
                                if (prev_param.get('pv') != expected_param['pv'] or
                                    prev_param.get('els') != expected_param['els'] or
                                    prev_param.get('element_type') != expected_param['element_type']):
                                    prev_match = False
                                    break
                                    
                                # Check appropriate pressure field based on stage
                                if stage_num == 1:
                                    if prev_param.get('feed_pressure') != expected_param['feed_pressure']:
                                        prev_match = False
                                        break
                                else:
                                    if prev_param.get('target_pressure') != expected_param['target_pressure']:
                                        prev_match = False
                                        break
                                    # Check boost pressure if provided
                                    if 'boost_pressure' in expected_param and (
                                        'boost_pressure' not in prev_param or 
                                        prev_param.get('boost_pressure') != expected_param['boost_pressure']):
                                        prev_match = False
                                        break
                        
                        if not prev_match:
                            continue
                    
                    matched_row = row
                    break
            
            if matched_row is None:
                pressure_field = 'feed_pressure' if stage == 1 else 'target_pressure'
                logger.error(f"No matching report found for stage {stage}, PV {pv}, ELS {els}, {pressure_field}: {pressure}")
                return None, None
            
            # Get the filename from the matched row
            filename = matched_row['filename']
            logger.info(f"Found matching report: {filename}")
            
            # Now read the element_flow sheet to get the feed pressure
            element_flow_df = pd.read_excel(self.prev_stage_excel_file, sheet_name='element_flow')
            
            # Find the row containing the filename
            matching_rows = element_flow_df.apply(
                lambda r: any(str(filename) in str(cell) for cell in r), axis=1
            )
            
            if not matching_rows.any():
                logger.error(f"No row found containing filename {filename} in element_flow sheet")
                return None,None
                
            row_idx = matching_rows[matching_rows].index[0]
            
            # Find columns that match the pattern 'row_*_Feed_Press'
            feed_press_columns = []
            pattern = re.compile(r'row_(\d+)_Feed_Press')

            for col in element_flow_df.columns:
                if pattern.match(col):
                    feed_press_columns.append(col)

            if not feed_press_columns:
                logger.error("No Feed_Press columns found in element_flow sheet")
                return None, None

            # Sort columns by the numeric part
            feed_press_columns.sort(key=lambda x: int(re.match(r'row_(\d+)_Feed_Press', x).group(1)))
            column_name = feed_press_columns[-1]

            logger.info(f"Using last Feed_Press column: {column_name}")

            # Get the cell value at the found row and specified column and round to 1 decimal
            feed_pressure = element_flow_df.loc[row_idx, column_name]
            rounded_feed_pressure = round(float(feed_pressure), 1)
            logger.info(f"Found feed pressure {rounded_feed_pressure} for stage {stage}")

            # get the boost pressure
            boost_pressure = matched_row['current_stage_params'].get("boost_pressure", None)
            return boost_pressure, rounded_feed_pressure
                
        except Exception as e:
            logger.error(f"Error reading Excel file: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return None, None

    def get_valid_target_pressures(self, target_values, feed_pressure):
        """
        Get valid target pressures based on feed pressure.
        
        Args:
            target_values (list): List of target pressures to check
            feed_pressure (float): Feed pressure to compare against
            
        Returns:
            list: List of valid target pressures
        """
        valid_pressures = []
        for target in target_values:
            boost = target - feed_pressure
            if boost > 0:
                valid_pressures.append(target)
        return valid_pressures
    
    def run_parameter_sweep(self, stage_configs):
        """
        Runs parameter sweep considering all previous stages.
        
        Args:
            stage_configs (list): List of stage configurations containing ranges
            
        Returns:
            tuple: (successful_runs, total_combinations)
        """
        successful_runs = 0
        total_combinations = 0
        restart_counter = 0
        # Number of successful runs before restarting the application
        restart_threshold = 50
        
        try:
            if self.stages == 1:
                # Stage 1 logic with step size
                config = stage_configs[0]
                pv_range = range(config['pv_range'][0], config['pv_range'][1] + 1)
                els_range = range(config['els_range'][0], config['els_range'][1] + 1)
                
                # Get pressure range with step size
                pressure_start = config['feed_pressure_range'][0]
                pressure_end = config['feed_pressure_range'][1]
                pressure_step = config.get('feed_pressure_step', 1)  # Default to 1 if not specified
                
                # Generate pressure values using step size
                pressure_values = [round(pressure_start + i * pressure_step, 1) 
                                   for i in range(int((pressure_end - pressure_start) / pressure_step) + 1)]
                
                element_type = config['element_type']
                total_combinations = len(pv_range) * len(els_range) * len(pressure_values)
                logger.info(f"Starting stage 1 parameter sweep with {total_combinations} combinations")
                logger.info(f"Pressure values: {pressure_values}")
                
                # Launch WAVE initially
                if not self.launch_wave():
                    logger.error("Failed to launch WAVE. Aborting parameter sweep.")
                    return 0, 0
                    
                for pv, els, pressure in itertools.product(pv_range, els_range, pressure_values):
                    cur_params = {
                        'stage': self.stages,
                        'pv': pv, 
                        'els': els,
                        'element_type': element_type,
                        'feed_pressure': pressure
                    }
                    
                    # Check if we need to restart the application
                    if restart_counter >= restart_threshold:
                        logger.info(f"Restarting WAVE application after {restart_counter} successful runs")
                        self.close_application()
                        time.sleep(5)  # Give the application time to close completely
                        if not self.launch_wave():
                            logger.error("Failed to restart WAVE. Aborting remaining parameter sweep.")
                            break
                        restart_counter = 0
                    
                    if self._process_parameter_combination(cur_params):
                        successful_runs += 1
                        restart_counter += 1
                        
            else:
                # Multi-stage logic (Stage 2 or 3)
                current_stage = self.stages
                current_config = stage_configs[current_stage - 1]
                
                # Get ranges for the current stage
                current_pv_range = range(current_config['pv_range'][0], current_config['pv_range'][1] + 1)
                current_els_range = range(current_config['els_range'][0], current_config['els_range'][1] + 1)
                
                # Get target pressure range with step size
                target_start = current_config['target_pressure_range'][0]
                target_end = current_config['target_pressure_range'][1]
                target_step = current_config.get('target_pressure_step', 1)  # Default to 1 if not specified
                
                # Generate target pressure values using step size
                target_values = [round(target_start + i * target_step, 1) 
                                 for i in range(int((target_end - target_start) / target_step) + 1)]
                
                cur_element_type = current_config['element_type']
                
                # Generate all combinations of previous stage parameters
                prev_stage_combinations = []
                for stage in range(1, current_stage):
                    config = stage_configs[stage - 1]
                    pv_range = range(config['pv_range'][0], config['pv_range'][1] + 1)
                    els_range = range(config['els_range'][0], config['els_range'][1] + 1)
                    element_type = config['element_type']
                    
                    if stage == 1:
                        # Use feed pressure step size for stage 1
                        pressure_start = config['feed_pressure_range'][0]
                        pressure_end = config['feed_pressure_range'][1]
                        pressure_step = config.get('feed_pressure_step', 1)
                        
                        pressure_values = [round(pressure_start + i * pressure_step, 1) 
                                          for i in range(int((pressure_end - pressure_start) / pressure_step) + 1)]
                    else:
                        # Use target pressure step size for stages > 1
                        pressure_start = config['target_pressure_range'][0]
                        pressure_end = config['target_pressure_range'][1]
                        pressure_step = config.get('target_pressure_step', 1)
                        
                        pressure_values = [round(pressure_start + i * pressure_step, 1) 
                                          for i in range(int((pressure_end - pressure_start) / pressure_step) + 1)]
                    
                    stage_params = list(itertools.product(pv_range, els_range, (element_type,), pressure_values))
                    prev_stage_combinations.append(stage_params)
                
                # Launch WAVE initially
                if not self.launch_wave():
                    logger.error("Failed to launch WAVE. Aborting parameter sweep.")
                    return 0, 0
                    
                # Process all combinations
                for prev_params in itertools.product(*prev_stage_combinations):
                    # Get feed pressure from previous stage's Excel
                    prev_stage = current_stage - 1
                    prev_pv, prev_els, prev_ele_type, prev_pressure = prev_params[prev_stage - 1]
                    
                    # Convert params to the format expected by get_feed_pressure_from_excel
                    prev_stage_tuples = []
                    for i, param_tuple in enumerate(prev_params):
                        stage_num = i + 1
                        if stage_num < prev_stage:
                            prev_stage_tuples.append(param_tuple)
                    
                    if prev_stage > 1:
                        prev_boost_pressure, feed_pressure = self.get_pressure_from_excel(
                            prev_stage, prev_pv, prev_els, prev_ele_type, prev_pressure, prev_stage_tuples)
                    else:
                        prev_boost_pressure, feed_pressure = self.get_pressure_from_excel(
                            prev_stage, prev_pv, prev_els, prev_ele_type, prev_pressure)
                         
                    if feed_pressure is None:
                        pressure_type = "feed_pressure" if prev_stage == 1 else "boost_pressure"
                        logger.info(f"Unable to find the previous stage results for the combination pv:{prev_pv} els:{prev_els} {pressure_type}: {prev_pressure}")
                        continue

                    # If boost pressure was found for stage > 1, update the tuple to include it
                    if prev_boost_pressure is not None and prev_stage > 1:
                        # Convert the tuple to list for modification
                        prev_param_list = list(prev_params[prev_stage - 1])
                        # Add boost pressure as the 5th element
                        if len(prev_param_list) == 4:  # Only add if not already there
                            prev_param_list.append(prev_boost_pressure)
                        
                        # Create new tuples list with the updated tuple
                        new_prev_params = list(prev_params)
                        new_prev_params[prev_stage - 1] = tuple(prev_param_list)
                        prev_params = tuple(new_prev_params)

                    # Get valid target pressures
                    valid_targets = self.get_valid_target_pressures(target_values, feed_pressure)
                    
                    if not valid_targets:
                        logger.info(f"No valid target pressures for prev_stage params: {prev_params}")
                        continue
                    
                    # Process current stage combinations
                    for target in valid_targets:
                        boost_pressure = round((target - feed_pressure), 1)
                        
                        for pv, els in itertools.product(current_pv_range, current_els_range):
                            total_combinations += 1
                            
                            # Check if we need to restart the application
                            if restart_counter >= restart_threshold:
                                logger.info(f"Restarting WAVE application after {restart_counter} successful runs")
                                self.close_application()
                                time.sleep(5)  # Give the application time to close completely
                                if not self.launch_wave():
                                    logger.error("Failed to restart WAVE. Aborting remaining parameter sweep.")
                                    return successful_runs, total_combinations
                                restart_counter = 0
                            
                            # Log the combination being processed
                            prev_stages_info = []
                            for i, p in enumerate(prev_params):
                                stage_num = i + 1
                                pressure_type = "feed_pressure" if stage_num == 1 else "boost_pressure"
                                prev_stages_info.append(
                                    f"Stage{stage_num}(PV={p[0]},ELS={p[1]},Ele_type={p[2]},{pressure_type}={p[3]})"
                                )
                            logger.info(f"Processing: {', '.join(prev_stages_info)}")
                            logger.info(f"Current Stage{current_stage}: PV={pv}, ELS={els}, Element_Type={cur_element_type}, Target={target}, Boost={boost_pressure}")
                            
                            cur_params = {
                                'stage': current_stage,
                                'pv': pv, 
                                'els': els,
                                'element_type': cur_element_type,
                                'boost_pressure': boost_pressure,
                                'target_pressure': target
                            }

                            if self._process_parameter_combination(cur_params, prev_params):
                                successful_runs += 1
                                restart_counter += 1
            
        finally:
            # Make sure to close the application before exiting
            self.close_application()
            
        logger.info(f"Parameter sweep completed. Successful runs: {successful_runs}/{total_combinations}")
        return successful_runs, total_combinations

    def _process_parameter_combination(self, cur_stage_params, prev_stage_params=None):
        """
        Process a single parameter combination.

        Args:
            cur_stage_params (dict): A dict of current stage parameters
            prev_stage_params (list, optional): List of tuples (pv, els, element_type, pressure) for each previous stage
            
        Returns:
            bool: True if successful, False otherwise
        """
        # Switch to Reverse Osmosis tab and set parameters
        if not self.select_tab(TAB_REVERSE_OSMOSIS):
            logger.error("Failed to select Reverse Osmosis tab")
            return False
        
        # Determine which pressure parameter to use based on stage
        stage = cur_stage_params['stage']
        pressure_param = cur_stage_params.get('feed_pressure' if stage == 1 else 'boost_pressure')
        
        if not self.set_reverse_osmosis_parameters(
            cur_stage_params['pv'], 
            cur_stage_params['els'], 
            cur_stage_params['element_type'], 
            pressure_param, 
            prev_stage_params):
            logger.error("Failed to set parameters")
            return False
            
        # Switch to Summary Report tab and open detailed report
        if not self.select_tab(TAB_SUMMARY_REPORT, timeout=10):
            logger.error("Failed to select Summary Report tab")
            return False
            
        if not self.open_detailed_report():
            logger.error("Failed to open detailed report")
            return False
                
        # Generate filename and add metadata
        filename = self.add_metadata_entry(cur_stage_params, prev_stage_params)
        actual_file_path = os.path.abspath(os.path.join(self.export_dir, filename))
        
        # Export to Excel
        if not self.export_to_excel(actual_file_path):
            logger.error("Failed to export report")
            return False
            
        return True
