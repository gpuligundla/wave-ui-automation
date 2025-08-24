# Configuration constants for WAVE application
import os

# Window titles and patterns
WINDOW_TITLE_PATTERN = "{project_name} - {case_name}"
CR_FLOW_WINDOW = "Reverse Osmosis Flow Calculator"
# Tab names
TAB_REVERSE_OSMOSIS = "Reverse Osmosis"
TAB_SUMMARY_REPORT = "Summary Report"

# UI element IDs and properties
UI_FIELDS = {
    "stages": {
        "type": "radio",
        "group_name": "Number of Stages",
        "class_name": "GroupBox"
    },
    "pv_per_stage": {
        "type": "edit",
        "auto_id": "txtStr1",
        "class_name": "TextBox"
    },
    "els_per_pv": {
        "type": "edit",
        "auto_id": "txtStr2",
        "class_name": "TextBox"
    },
    "feed_pressure": {
        "type": "edit",
        "auto_id": "FeedPressureRO",
        "class_name": "TextBox"
    },
    "boost_pressure": {
        "type": "edit",
        "auto_id": "txtStr5",
        "class_name": "TextBox"
    },
    "feed_flow_rate": {
        "type": "edit",
        "class_name": "TextBox",
        "group_name": "Flows",
        "group_class_name": "GroupBox"
    },
    "element_type": {
        "type": "combo",
        "auto_id": ["0", "1", "2"], 
        "class_name": "ComboBox"
    },
    "detailed_report": {
        "type": "Button",
        "class_name":"RibbonButton"
    },
    "file_name": {
        "type": "edit",
        "class_name": "Edit",
        "auto_id": "1001"
    },
    "save_button": {
        "type": "button",
        "class_name": "Button",
        "auto_id": "1"
    },
    "cr_flow_edit": {
        "type": "edit",
        "class_name": "TextBox",
        "group_name": "Flows",
        "group_class_name": "GroupBox"
    },
    "cr_flow": {
        "type": "edit",
        "class_name": "TextBox",
        "group_name": "Pass 1",
        "group_class_name": "GroupBox",
        "auto_id": "txtConcRecyclePercent"
    }
}


# Button and control names
BUTTON_DETAILED_REPORT = "Detailed Report"
BUTTON_EXPORT = "Export"
TOOLBAR_ID = "toolStrip1"

# Dialog information
SAVE_DIALOG_CLASS = "#32770"
SAVE_DIALOG_TITLE = "Save As"
CONFIRM_DIALOG_TITLE = "Confirm Save As"
EXIT_DIALOG_TITLE = "Confirmation"

# Default parameters for simulation
DEFAULT_PARAMETERS = {
    "feed_flow_rate": 2.1,
    "stages": 1,
    "pv_per_stage": 1,
    "els_per_pv": 1,
    "element_type": "NF90-4040",
    "feed_pressure": 10,
    "target_pressure": 10,
    "pressure_factor": 1
}

# Parameter ranges for sweep
PARAMETER_RANGES = {
    "pv_per_stage": range(1, 11),  # 1-10
    "els_per_pv": range(1, 9),    # 1-8
    "feed_pressure": (10, 50)    # 10-40
}

# Wait times for various operations (in seconds)
WAIT_TIMES = {
    "app_load": 15,
    "tab_switch": 5,
    "input_delay": 1,
    "report_gen": 10,
    "dialog_wait": 10,
    "save_complete": 2
}

APPLICATION_LOAD_TIMEOUT = 10
REPORT_GENERATION_TIMEOUT = 10
WINDOW_VISIBLE_TIMEOUT = 2
TAB_SWITCH_TIMEOUT = 5
DEFAULT_TIMEOUT = 5
SAVE_FILE_TIMEOUT = 2

# Export settings
EXPORT_DIR = os.path.join(os.path.expanduser("~"), "Downloads", "WAVE_Reports")