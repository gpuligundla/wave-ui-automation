"""
Main script for running the WAVE UI automation.
"""
import argparse
import logging
import os
from datetime import datetime
from waveui import WaveUI
from constants import DEFAULT_PARAMETERS, PARAMETER_RANGES

logger = logging.getLogger(__name__)

def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(description='WAVE UI Automation')
    
    parser.add_argument('--wave-exe', required=True,
                        help='Path to the WAVE executable')
                        
    parser.add_argument('--project-path', required=True,
                        help='Path to the WAVE project file')
    
    parser.add_argument('--project-name', required=True,
                        help='Path to the WAVE project file')
                        
    parser.add_argument('--case-name', required=True,
                        help='Name of the case in the WAVE project')
                        
    parser.add_argument('--stages', type=int, default=DEFAULT_PARAMETERS['stages'],
                        help='Number of stages')
    
    parser.add_argument('--feed-flow-rate', type=float, default=DEFAULT_PARAMETERS['feed_flow_rate'],
                        help='Feed flow rate value')
                        
    parser.add_argument('--pv-min', type=int, default=min(PARAMETER_RANGES['pv_per_stage']),
                        help='Minimum PV per stage value')
                        
    parser.add_argument('--pv-max', type=int, default=max(PARAMETER_RANGES['pv_per_stage']),
                        help='Maximum PV per stage value')
                        
    parser.add_argument('--els-min', type=int, default=min(PARAMETER_RANGES['els_per_pv']),
                        help='Minimum elements per PV value')
                        
    parser.add_argument('--els-max', type=int, default=max(PARAMETER_RANGES['els_per_pv']),
                        help='Maximum elements per PV value')
    
    parser.add_argument('--element-type', default=DEFAULT_PARAMETERS['element_type'],
                        help='Element type')
    
    parser.add_argument('--feed-pressure-min', type=int, default=PARAMETER_RANGES['feed_pressure'],
                        help='Minimum Feed pressure value to test')                 
    
    parser.add_argument('--feed-pressure-max', type=int, default=PARAMETER_RANGES['feed_pressure'],
                        help='Maximum Feed pressure value to test') 
    
    return parser.parse_args()

def display_start_messge(args):
    """Log the start of automation with given parameters."""
    logger.info(f"WAVE UI Automation - Started")
    logger.info(f"WAVE executable: {args.wave_exe}")
    logger.info(f"Project file: {args.project_path}")
    logger.info(f"Project name: {args.project_name}")
    logger.info(f"Case name: {args.case_name}")
    logger.info(f"Parameter ranges:")
    logger.info(f"PV per stage: {args.pv_min} to {args.pv_max}")
    logger.info(f"Elements per PV: {args.els_min} to {args.els_max}")
    logger.info(f"Feed pressures: {args.feed_pressure_min} to {args.feed_pressure_max}")
    

def main():
    """Main function to run the WAVE automation."""
    # Parse command line arguments
    args = parse_arguments()
    
    time = datetime.now()
    str_time = time.strftime("%Y%m%d-%H%M%S")
    
    # Ensure logs directory exists
    logs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    if not os.path.exists(logs_dir):
        os.makedirs(logs_dir)
        print(f"Created logs directory: {logs_dir}")
    
    logging.basicConfig(filename=os.path.join(logs_dir, f"{str_time}-wave-ui-automation.log"), 
                        level=logging.INFO,
                        format='[%(asctime)s] - [%(levelname)s] - %(message)s',datefmt='%Y-%m-%d %H:%M:%S')

    # Log the start of automation
    display_start_messge(args)
    

    pv_range = range(args.pv_min, args.pv_max + 1)
    els_range = range(args.els_min, args.els_max + 1)
    pressure_range = range(args.feed_pressure_min, args.feed_pressure_max + 10, 10)
    
    # Create WaveUI instance
    wave_ui = WaveUI(
        file_name=args.wave_exe,
        project_path=args.project_path,
        project_name=args.project_name,
        case_name=args.case_name,
        feed_flow_rate=args.feed_flow_rate,
        stages=args.stages,
        element_type=args.element_type
    )
    
    # Run parameter sweep with the specified ranges
    successful_runs, total_combinations = wave_ui.run_parameter_sweep(
        pv_range=pv_range,
        els_range=els_range,
        pressure_range=pressure_range
    )
    
    # Final report
    success_rate = (successful_runs / total_combinations) * 100 if total_combinations > 0 else 0
    logger.info(f"Automation completed at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"Total combinations: {total_combinations}")
    logger.info(f"Successful runs: {successful_runs} ({success_rate:.1f}%)")

if __name__ == "__main__":
    main()