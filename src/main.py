"""
Main script for running the WAVE UI automation.
"""
import logging
import os
import json
import sys
from datetime import datetime
from waveui import WaveUI
from utils import load_json_config

logger = logging.getLogger(__name__)

def setup_logging():
    """Set up logging configuration."""
    time = datetime.now()
    str_time = time.strftime("%Y%m%d-%H%M%S")
    
    logs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    if not os.path.exists(logs_dir):
        os.makedirs(logs_dir)
        print(f"Created logs directory: {logs_dir}")
    
    logging.basicConfig(
        filename=os.path.join(logs_dir, f"{str_time}-wave-ui-automation.log"), 
        level=logging.INFO,
        format='[%(asctime)s] - [%(levelname)s] - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

def run_wave_automation(config_file, stage_to_run, prev_stage_excel=None, add_conc_recycle=False):
    """Run the automation using a JSON configuration file for a specific stage."""
    try:
        # Load and validate the configuration
        config = load_json_config(config_file)
        logger.info(f"Using JSON configuration from: {config_file}")
        
        # Validate stage number
        if stage_to_run < 1 or stage_to_run > len(config['stages']):
            raise ValueError(f"Invalid stage number. Must be between 1 and {len(config['stages'])}")
        
        # If stage > 1, require previous stage excel file
        if stage_to_run > 1 and not prev_stage_excel:
            raise ValueError("Previous stage Excel file path is required for stages > 1")
        
        # Extract common parameters
        common = config['common']
        wave_exe = common['wave_exe']
        project_path = common['project_path']
        project_name = common['project_name']
        case_name = common['case_name']
        feed_flow_rate = common['feed_flow_rate']
        export_dir = common['export_dir']
        # Get configurations for current and previous stages
        stage_configs = config['stages'][:stage_to_run]
        conc_recycle_flow = config['optional']['conc_recycle_flow'] if add_conc_recycle else []

        # Create WaveUI instance
        wave_ui = WaveUI(
            file_name=wave_exe,
            project_path=project_path,
            project_name=project_name,
            case_name=case_name,
            config=config,
            feed_flow_rate=feed_flow_rate,
            stages=stage_to_run,
            prev_stage_excel_file=prev_stage_excel,
            conc_recycle_flow = conc_recycle_flow,
            export_dir=export_dir
        )
        
        # Run parameter sweep with all stage configurations
        successful_runs, total_combinations = wave_ui.run_parameter_sweep(stage_configs)
        
        # Report results
        success_rate = (successful_runs / total_combinations) * 100 if total_combinations > 0 else 0
        logger.info(f"Stage {stage_to_run} completed: {successful_runs}/{total_combinations} successful runs ({success_rate:.1f}%)")
        
    except Exception as e:
        logger.error(f"Error running WAVE automation: {e}")
        raise

def main():
    """Main function to run the WAVE automation."""
    try:
        # Setup logging
        setup_logging()
        
        # Parse command line arguments
        if len(sys.argv) < 3:
            print("Usage: python main.py <config_file> <stage_number> [prev_stage_excel] [--conc-recycle]")
            print("Example: python main.py config.json 1 --conc-recycle")
            print("Example: python main.py config.json 2 reports/stage1_results.xls")
            sys.exit(1)

        config_file = sys.argv[1]
        if not os.path.exists(config_file):
            print(f"Error: Configuration file '{config_file}' not found")
            sys.exit(1)

        try:
            stage_to_run = int(sys.argv[2])
        except ValueError:
            print("Error: Stage number must be an integer")
            sys.exit(1)

        prev_stage_excel = None
        add_conc_recycle = False

        # Parse optional arguments
        args = sys.argv[3:]
        for arg in args:
            if arg == "--conc-recycle":
                add_conc_recycle = True
            elif prev_stage_excel is None:
                prev_stage_excel = arg
                if not os.path.exists(prev_stage_excel):
                    print(f"Error: Previous stage Excel file '{prev_stage_excel}' not found")
                    sys.exit(1)

        # Run the automation
        run_wave_automation(config_file, stage_to_run, prev_stage_excel, add_conc_recycle)
        
        logger.info(f"Automation completed at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
    except Exception as e:
        logger.error(f"Fatal error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()