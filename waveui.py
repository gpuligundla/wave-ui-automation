"""
This class is responsible for WAVE UI automation actions
"""
from pywinauto import Application, Desktop, keyboard
import os
import time
import itertools
import logging
import pandas as pd

from constants import (
    WINDOW_TITLE_PATTERN, TAB_REVERSE_OSMOSIS, TAB_SUMMARY_REPORT, 
    UI_FIELDS, BUTTON_DETAILED_REPORT, BUTTON_EXPORT, 
    TOOLBAR_ID, EXPORT_DIR, EXPORT_FILENAME_TEMPLATE,
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
                 prev_stage_excel_file=None):
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
        
        self.export_dir = check_and_create_results_directory(EXPORT_DIR)
    
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
                logger.info(f"Previous stage(s) values are {prev_stage_params} of format (pv, els, pressure)")
                cur_params = (pv_per_stage, els_per_pv, element_type, pressure)
                params = prev_stage_params + (cur_params,) 
                for cur_stage, (pv, els, ele_type, pressure) in enumerate(params, start=1):
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

        # 6. Handle Element Type (combo box)
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
        Handles the exit confirmation dialog.
        
        Returns:
            bool: True if successful, False otherwise.
        """
        try:
            # Wait for the confirmation dialog
            confirm_dialog = Desktop(backend="win32").window(class_name=SAVE_DIALOG_CLASS, title=EXIT_DIALOG_TITLE)
            confirm_dialog.wait('exists', DEFAULT_TIMEOUT)
            
            # Find the Yes button using the exact automation ID and class name
            yes_button = confirm_dialog.child_window(title="Yes", class_name="Button", auto_id="6")
            
            if yes_button:
                yes_button.click_input()
                logger.info("Confirmed application exit")
                return True
            else:
                logger.error("Could not find Yes button in exit confirmation dialog")
                return False
        except Exception as e:
            logger.error(f"Error handling exit confirmation: {e}")
            return False
    
    def get_feed_pressure_from_excel(self, stage, pv, els, pressure):
        """
        Get feed pressure from a specific stage's Excel report.
        
        Args:
            stage (int): Stage number of the Excel report
            pv (int): PV value used in filename
            els (int): ELS value to look up in excel
            pressure (float): Pressure value used in filename
            
        Returns:
            float: Feed pressure value rounded to 1 decimal place if found, None otherwise
        """
        try:
            # Construct filename for the Excel report
            filename = f"{EXPORT_FILENAME_TEMPLATE.format(stage=stage, pv=pv, els=els, pressure=pressure)}.xls"
            
            # Read the element_flow sheet
            df = pd.read_excel(self.prev_stage_excel_file, sheet_name='element_flow')
            
            # Find the row where the filename matches a cell value
            matching_rows = df.apply(lambda row: row.astype(str).str.contains(filename).any(), axis=1)
            if not matching_rows.any():
                logger.error(f"No row found containing filename {filename}")
                return None
                
            row_idx = matching_rows[matching_rows].index[0]
            
            # Construct column name for the feed pressure
            column_name = f'row_{els}_Feed_Press'
            
            # Get the cell value at the found row and specified column and round to 1 decimal
            if column_name in df.columns:
                feed_pressure = df.loc[row_idx, column_name]
                rounded_pressure = round(float(feed_pressure), 1)
                logger.info(f"Found feed pressure {rounded_pressure} for stage {stage}, ELS {els}")
                return rounded_pressure
            else:
                logger.error(f"Column {column_name} not found in Excel file")
                return None
                
        except Exception as e:
            logger.error(f"Error reading Excel file: {e}")
            return None

    def get_valid_target_pressures(self, target_range, feed_pressure):
        """
        Get valid target pressures based on feed pressure.
        
        Args:
            target_range (range): Range of target pressures to check
            feed_pressure (float): Feed pressure to compare against
            
        Returns:
            list: List of valid target pressures
        """
        valid_pressures = []
        for target in target_range:
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
        if not self.launch_wave():
            logger.error("Failed to launch WAVE. Aborting parameter sweep.")
            return 0, 0
            
        successful_runs = 0
        total_combinations = 0
        
        try:
            if self.stages == 1:
                # Stage 1 logic remains the same
                config = stage_configs[0]
                pv_range = range(config['pv_range'][0], config['pv_range'][1] + 1)
                els_range = range(config['els_range'][0], config['els_range'][1] + 1)
                pressure_range = range(config['feed_pressure_range'][0], 
                                    config['feed_pressure_range'][1] + 1)
                element_type = config['element_type']
                total_combinations = len(pv_range) * len(els_range) * len(pressure_range)
                logger.info(f"Starting stage 1 parameter sweep with {total_combinations} combinations")
                
                for pv, els, pressure in itertools.product(pv_range, els_range, pressure_range):
                    filename = f"{EXPORT_FILENAME_TEMPLATE.format(stage=1, pv=pv, els=els, pressure=pressure)}.xls"
                    file_path = os.path.join(self.export_dir, filename)
                    
                    if self._process_parameter_combination(pv, els, element_type, pressure, file_path):
                        successful_runs += 1
                        
            else:
                # Multi-stage logic (Stage 2 or 3)
                current_stage = self.stages
                current_config = stage_configs[current_stage - 1]
                
                # Get ranges for the current stage
                current_pv_range = range(current_config['pv_range'][0], current_config['pv_range'][1] + 1)
                current_els_range = range(current_config['els_range'][0], current_config['els_range'][1] + 1)
                target_range = range(current_config['target_pressure_range'][0], 
                                   current_config['target_pressure_range'][1] + 1)
                cur_element_type = current_config['element_type']
                # Generate all combinations of previous stage parameters
                prev_stage_combinations = []
                for stage in range(1, current_stage):
                    config = stage_configs[stage - 1]
                    pv_range = range(config['pv_range'][0], config['pv_range'][1] + 1)
                    els_range = range(config['els_range'][0], config['els_range'][1] + 1)
                    element_type = config['element_type']
                    
                    if stage == 1:
                        pressure_range = range(config['feed_pressure_range'][0], 
                                            config['feed_pressure_range'][1] + 1)
                    else:
                        pressure_range = range(config['target_pressure_range'][0], 
                                            config['target_pressure_range'][1] + 1)
                    
                    stage_params = list(itertools.product(pv_range, els_range, (element_type,), pressure_range))
                    prev_stage_combinations.append(stage_params)
                
                # Process all combinations
                for prev_params in itertools.product(*prev_stage_combinations):
                    # Get feed pressure from previous stage's Excel
                    prev_stage = current_stage - 1
                    prev_pv, prev_els, _, prev_pressure = prev_params[prev_stage - 1]
                    
                    feed_pressure = self.get_feed_pressure_from_excel(
                        prev_stage, prev_pv, prev_els, prev_pressure)
                    
                    if feed_pressure is None:
                        logger.info(f"Unable to find the previous stage results for the combination pv:{prev_pv} els:{prev_els} feedpressure: {prev_pressure}")
                        continue
                    
                    # Get valid target pressures
                    valid_targets = self.get_valid_target_pressures(target_range, feed_pressure)
                    
                    if not valid_targets:
                        logger.info(f"No valid target pressures for prev_stage params: {prev_params}")
                        continue
                    
                    # Process current stage combinations
                    for target in valid_targets:
                        boost_pressure = target - feed_pressure
                        
                        for pv, els in itertools.product(current_pv_range, current_els_range):
                            total_combinations += 1
                            
                            filename = f"{EXPORT_FILENAME_TEMPLATE.format(stage=current_stage, pv=pv, els=els, pressure=target)}.xls"
                            file_path = os.path.join(self.export_dir, filename)
                            
                            # Log the combination being processed
                            prev_stages_info = ", ".join(
                                f"Stage{i+1}(PV={p[0]},ELS={p[1]},Ele_tye={p[2]}, P={p[3]})" 
                                for i, p in enumerate(prev_params)
                            )
                            logger.info(f"Processing: {prev_stages_info}")
                            logger.info(f"Current Stage{current_stage}: PV={pv}, ELS={els}, Element_Type={cur_element_type}, Target={target}, Boost={boost_pressure}")
                            
                            if self._process_parameter_combination(pv, els, cur_element_type, boost_pressure, file_path, prev_params):
                                successful_runs += 1
                
        finally:
            self.close_application()
            
        logger.info(f"Parameter sweep completed. Successful runs: {successful_runs}/{total_combinations}")
        return successful_runs, total_combinations

    def _process_parameter_combination(self, pv, els, element_type, pressure, file_path, prev_stage_params=None):
        """
        Process a single parameter combination.
        
        Args:
            pv (int): PV value for the current stage
            els (int): ELS value for the current stage
            element_type (str): Element type of the current stage
            pressure (float): Pressure value for the current stage (feed pressure for stage 1, boost pressure for stages 2-3)
            file_path (str): Path to save the Excel report
            prev_stage_params (list, optional): List of tuples (pv, els, pressure) for each previous stage. Required for stages > 1.
            
        Returns:
            bool: True if successful, False otherwise
        """
        # Switch to Reverse Osmosis tab and set parameters
        if not self.select_tab(TAB_REVERSE_OSMOSIS):
            logger.error("Failed to select Reverse Osmosis tab")
            return False
        
        if not self.set_reverse_osmosis_parameters(pv, els, element_type, pressure, prev_stage_params):
            logger.error("Failed to set parameters")
            return False
            
        # Switch to Summary Report tab and open detailed report
        if not self.select_tab(TAB_SUMMARY_REPORT, timeout=10):
            logger.error("Failed to select Summary Report tab")
            return False
            
        if not self.open_detailed_report():
            logger.error("Failed to open detailed report")
            return False
            
        # Export to Excel
        if not self.export_to_excel(file_path):
            logger.error("Failed to export report")
            return False
            
        return True