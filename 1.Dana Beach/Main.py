import subprocess
import psutil
import sys
import time
import logging
import threading
from pathlib import Path
import os
import pyautogui
import ctypes
from datetime import datetime, timedelta
pyautogui.FAILSAFE = False

yesterday = datetime.now() - timedelta(days=1)
formatted_date = yesterday.strftime('%d%m%y')

# Disable Caps Lock if active
if ctypes.WinDLL('User32.dll').GetKeyState(0x14) & 1:
    pyautogui.press('capslock')

# Dummy credentials and dates (not used directly in this script)
username = "a.ibrahim"
password = "Des101"
start_date = "011124"
end_date = "011125"

# Logging setup
logging.basicConfig(
    filename="Report_Automation.log",
    level=logging.INFO,
    format="%(asctime)s - %(message)s"
)

def delete_WB_xls_files():
    folder_path = Path(r"D:\Daily_Chart\Chart_Resources\S25\HRG")
    try:
        for file in folder_path.glob("WB*"):
            try:
                file.unlink()
                logging.info(f"Deleted file: {file}")
                print(f"Deleted file: {file}")
            except Exception as e:
                logging.error(f"Failed to delete {file}: {e}")
    except Exception as e:
        logging.error(f"Error accessing folder: {e}")
        sys.exit(1)

def monitor_fidelio_process():
    try:
        while True:
            fidelio_running = False
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'] and proc.info['name'].lower() == 'fideliov8.exe':
                    fidelio_running = True
                    break

            if not fidelio_running:
                logging.info("FidelioV8.exe process not found - terminating program")
                print("FidelioV8.exe process not found - terminating program")
                sys.exit(0)

            time.sleep(1)
    except Exception as e:
        logging.error(f"Error monitoring Fidelio process: {str(e)}")
        sys.exit(1)

def run_scripts():
    current_dir = Path(__file__).parent
    folder_path = Path(r"D:\Daily_Chart\Chart_Resources\S25\HRG")

    script_file_map = {
        "C.RTA.py":  "DB RTA.xls",
        "D.RTO.py":  "DB RTO.xls",
        "E.RCT.py":  "DB RCT.xls",
        "E.RCR.py":  "DB RCR.xls",
        "F.AR.py" :  "DB AR.xls",
        "C.RM.py" :  "DB RM.xls",
        "C.RC.py" :  "DB RC.xls",
        "H&F.py"  :  "DB H&F.xlsx",
        "CF.py"   : "DB CF.xlsx"
    }

    def run_script(script):
        try:
            script_path = current_dir / script
            result = subprocess.run(["python", str(script_path)], check=True)
            return result.returncode == 0
        except subprocess.CalledProcessError as e:
            logging.error(f"Error running {script}: {e}")
            return False
        except FileNotFoundError:
            logging.error(f"Script not found: {script}")
            return False

    def check_missing_files():
        return [
            script for script, expected_file in script_file_map.items()
            if not (folder_path / expected_file).exists()
        ]

    def run_launch_and_scripts(scripts_to_run, start_monitor=False):
        if run_script("B.Launch.py"):
            if start_monitor:
                threading.Thread(target=monitor_fidelio_process, daemon=True).start()
            for script in scripts_to_run:
                run_script(script)
            return True
        else:
            logging.error("Failed to run B.Launch.py")
            return False

    # Step 1: Cleanup
    delete_WB_xls_files()

    try:
        # Step 2: Run all scripts initially
        run_launch_and_scripts(script_file_map.keys(), start_monitor=True)

        # Step 3: Retry logic if outputs are missing
        max_retries = 3
        retry_count = 0

        while retry_count < max_retries:
            missing_scripts = check_missing_files()
            if not missing_scripts:
                logging.info("All expected files created successfully.")
                break

            retry_count += 1
            logging.warning(f"Retry {retry_count}: Missing files: {missing_scripts}")
            logging.info("Re-running Launch and missing scripts...")

            run_launch_and_scripts(missing_scripts, start_monitor=False)
            time.sleep(5)  # Optional wait between retries

        # Step 4: Final check
        final_missing = check_missing_files()
        if final_missing:
            logging.error(f"Files still missing after {max_retries} retries: {final_missing}")
        else:
            logging.info("All expected files created successfully after retries.")

        # Step 5: Kill Fidelio process
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] and proc.info['name'].lower() == "fideliov8.exe":
                proc.kill()
                logging.info("FidelioV8.exe process terminated")
                break

    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    run_scripts()
