"""
This class is responsible for WAVE UI automation actions
"""
from pywinauto import Application, Desktop, keyboard
import os
import time
import itertools
import logging

from constants import (
    WINDOW_TITLE_PATTERN, TAB_REVERSE_OSMOSIS, TAB_SUMMARY_REPORT, 
    UI_FIELDS, BUTTON_DETAILED_REPORT, BUTTON_EXPORT, 
    TOOLBAR_ID, EXPORT_DIR, EXPORT_FILENAME_TEMPLATE,
    EXIT_DIALOG_TITLE, APPLICATION_LOAD_TIMEOUT, REPORT_GENERATION_TIMEOUT, 
    TAB_SWITCH_TIMEOUT, DEFAULT_TIMEOUT, WINDOW_VISIBLE_TIMEOUT, SAVE_DIALOG_CLASS
)
from utils import check_and_create_results_directory, handle_save_dialog

logger = logging.getLogger(__name__)

class WaveUI:
    """
    Class for automating interactions with the DuPont WAVE software.
    """
    
    def __init__(self, file_name, project_path, project_name, case_name, feed_flow_rate=2.1, stages=1, 
                 pv_per_stage=1, els_per_pv=1, element_type="NF90-4040", feed_pressure=10, boost_pressure=10):
        """
        Initialize the WaveUI automation class.
        
        Args:
            file_name (str): Path to the WAVE executable
            project_path (str): Path to the project file
            project_name (str): Name of the project to work with
            case_name (str): Name of the case to work with
            feed_flow_rate (float): Feed flow rate value
            stages (int): Number of stages
            pv_per_stage (int): PV per stage value
            els_per_pv (int): Elements per PV value
            element_type (str): Element type string
            feed_pressure (float): Feed pressure value
            boost_pressure (float): Boost pressure value (used only when stages > 1)
        """
        self.file_name = file_name
        self.project_path = project_path
        self.project_name = project_name
        self.case_name = case_name
        self.feed_flow_rate = feed_flow_rate
        self.stages = stages
        self.pv_per_stage = pv_per_stage
        self.els_per_pv = els_per_pv
        self.element_type = element_type
        self.feed_pressure = feed_pressure
        self.boost_pressure = boost_pressure
        
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
    
    def select_tab(self, tab_name):
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
            tab.wait('enabled', timeout=TAB_SWITCH_TIMEOUT)
            return True
        except Exception as e:
            logger.error(f"Failed to select tab '{tab_name}': {e}")
            return False
    
    def set_reverse_osmosis_parameters(self, pv_per_stage, els_per_pv, feed_pressure):
        """
        Sets parameters in the Reverse Osmosis tab.
        
        Args:
            pv_per_stage (int): Number of pressure vessels per stage
            els_per_pv (int): Number of elements per pressure vessel
            feed_pressure (int): Feed pressure value
            
        Returns:
            bool: True if successful, False otherwise.
        """
        try:
            logger.info(f"Setting RO parameters: PV={pv_per_stage}, ELS={els_per_pv}, Pressure={feed_pressure}")
                        
            # 1. Handle Stages (radio buttons)
            try:
                stages_group = self.main_window.child_window(
                    title=UI_FIELDS["stages"]["group_name"],
                    control_type="Group",
                    class_name=UI_FIELDS["stages"]["class_name"]
                )
                # Find the radio button that matches our stages value
                radio_button = stages_group.child_window(
                    title=f"{self.stages}  ",  # Note the spaces in the radio button title
                    control_type="RadioButton"
                )
                radio_button.click_input()
                logger.info(f"Set stages to {self.stages}")
            except Exception as e:
                logger.error(f"Failed to set stages: {e}")
            
            # 2. Handle PV per stage (edit field)
            try:
                pv_edit = self.main_window.child_window(
                    auto_id=UI_FIELDS["pv_per_stage"]["auto_id"], 
                    control_type="Edit",
                    class_name=UI_FIELDS["pv_per_stage"]["class_name"]
                )
                pv_edit.set_text(str(pv_per_stage))
                logger.info(f"Set PV per stage to {pv_per_stage}")
            except Exception as e:
                logger.error(f"Failed to set PV per stage: {e}")
            
            # 3. Handle Elements per PV (edit field)
            try:
                els_edit = self.main_window.child_window(
                    auto_id=UI_FIELDS["els_per_pv"]["auto_id"], 
                    control_type="Edit",
                    class_name=UI_FIELDS["els_per_pv"]["class_name"]
                )
                els_edit.set_text(str(els_per_pv))
                logger.info(f"Set elements per PV to {els_per_pv}")
            except Exception as e:
                logger.error(f"Failed to set elements per PV: {e}")
            
            # 4. Handle Feed Pressure (edit field)
            try:
                # Use the specific AutomationId to find the feed pressure edit field
                feed_pressure_edit = self.main_window.child_window(
                    auto_id=UI_FIELDS["feed_pressure"]["auto_id"], 
                    control_type="Edit",
                    class_name=UI_FIELDS["feed_pressure"]["class_name"]
                )
                
                feed_pressure_edit.set_text(str(feed_pressure))
                logger.info(f"Set feed pressure to {feed_pressure}")
            except Exception as e:
                logger.error(f"Failed to set feed pressure: {e}")
            
            # 5. Handle Feed Flow Rate (edit field within Flows group)
            # try:
            #     # First find the Flows group
            #     flows_group = self.main_window.child_window(
            #         title=UI_FIELDS["feed_flow_rate"]["group_name"],
            #         control_type="Group",
            #         class_name=UI_FIELDS["feed_flow_rate"]["group_class_name"]
            #     )
                
            #     # Since there's no auto_id for feed_flow_rate, find the TextBox by its class name
            #     # This will find all TextBox controls in the Flows group
            #     textboxes = flows_group.children(class_name=UI_FIELDS["feed_flow_rate"]["class_name"])
                
            #     # Assuming the first TextBox in the Flows group is the feed flow rate control
            #     if textboxes:
            #         flow_rate_edit = textboxes[0]
            #         flow_rate_edit.set_text(str(self.feed_flow_rate))
            #         print(f"Set feed flow rate to {self.feed_flow_rate}")
            #     else:
            #         print("Could not find feed flow rate TextBox in Flows group")
            # except Exception as e:
            #     print(f"Failed to set feed flow rate: {e}")
            
            # 6. Handle Element Type (combo box)
            try:
                combo_box = self.main_window.child_window(
                    auto_id=UI_FIELDS["element_type"]["auto_id"],
                    control_type="ComboBox",
                    class_name=UI_FIELDS["element_type"]["class_name"]
                )
                combo_box.select(self.element_type)
                logger.info(f"Set element type to {self.element_type}")
            except Exception as e:
                logger.error(f"Failed to set element type: {e}")
        
            return True
        except Exception as e:
            logger.error(f"Error setting RO parameters: {e}")
            return False
    
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
            # time.sleep(REPORT_GENERATION_TIMEOUT)
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
    
    def run_parameter_sweep(self, pv_range, els_range, pressure_range):
        """
        Runs through all parameter combinations.
        
        Args:
            pv_range (range): Range of PV per stage values to test
            els_range (range): Range of elements per PV values to test
            pressure_range (range): Range of pressure values to test
            
        Returns:
            tuple: (successful_runs, total_combinations)
        """
        
        # Launch WAVE
        if not self.launch_wave():
            logger.error("Failed to launch WAVE. Aborting parameter sweep.")
            return 0, 0
        
        # Track successful runs
        successful_runs = 0
        total_combinations = len(pv_range) * len(els_range) * len(pressure_range)
        logger.info(f"Starting parameter sweep with {total_combinations} combinations")
        
        # Iterate through parameter combinations
        for pv, els, pressure in itertools.product(pv_range, els_range, pressure_range):
            # Generate file path and handle save dialog
            filename = EXPORT_FILENAME_TEMPLATE.format(pv=pv, els=els, pressure=pressure)
            file_path = os.path.join(self.export_dir, filename)

            try:
                logger.info(f"--- Testing combination: PV={pv}, ELS={els}, Pressure={pressure} ---")
                
                # Switch to Reverse Osmosis tab and set parameters
                if not self.select_tab(TAB_REVERSE_OSMOSIS):
                    logger.error("Failed to select Reverse Osmosis tab. Skipping combination.")
                    continue
                
                if not self.set_reverse_osmosis_parameters(pv, els, pressure):
                    logger.error("Failed to set parameters. Skipping combination.")
                    continue
                
                # Switch to Summary Report tab and open detailed report
                if not self.select_tab(TAB_SUMMARY_REPORT):
                    logger.error("Failed to select Summary Report tab. Skipping combination.")
                    continue
                
                if not self.open_detailed_report():
                    logger.error("Failed to open detailed report. Skipping combination.")
                    continue
                
                # Export to Excel
                if self.export_to_excel(file_path):
                    successful_runs += 1
                else:
                    logger.error("Failed to export report. Continuing with next combination.")
                
            except Exception as e:
                logger.error(f"Error processing combination: {e}")
        
        # Close the application when done
        self.close_application()
        
        logger.info(f"\nParameter sweep completed. Successful runs: {successful_runs}/{total_combinations}")
        return successful_runs, total_combinations