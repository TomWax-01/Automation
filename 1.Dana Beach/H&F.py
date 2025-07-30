import pyautogui
import time
import pygetwindow as gw
import logging
import os
import pywinauto
import psutil  # <-- added to handle killing Excel
from pywinauto.application import Application

from Main import start_date, end_date

# Setup logging
log_file = os.path.join(os.path.dirname(__file__), "Report_Automation.log")
logging.basicConfig(
    filename=log_file,
    filemode='a',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s')

def wait_for_window(title, max_wait_time=100):
    logging.info(f"Waiting for window: {title}")
    end_time = time.time() + max_wait_time
    while time.time() < end_time:
        windows = gw.getWindowsWithTitle(title)
        if windows:
            return windows[0]
        time.sleep(0.2)
    logging.error(f"Window '{title}' not found after {max_wait_time} seconds.")
    return None

def wait_for_window_to_stabilize(window, timeout=15):
    logging.info("Waiting for window to stabilize...")
    end_time = time.time() + timeout
    last_rect = (window.left, window.top, window.width, window.height)
    while time.time() < end_time:
        time.sleep(0.5)
        window_rect = (window.left, window.top, window.width, window.height)
        if window_rect == last_rect:
            logging.info("Window stabilized.")
            return True
        last_rect = window_rect
    logging.error("Window did not stabilize.")
    return False

def get_window_by_title(title):
    windows = gw.getWindowsWithTitle(title)
    return windows[0] if windows else None

def wait_for_excel_open(max_wait_time=60):
    logging.info("Waiting for Excel window to appear...")
    end_time = time.time() + max_wait_time
    while time.time() < end_time:
        windows = gw.getWindowsWithTitle("Excel")
        if windows:
            logging.info("Excel window detected.")
            return windows[0]
        time.sleep(1)
    logging.error("Excel window did not open in time.")
    return None

def save_excel_file(save_path, file_name):
    try:
        full_path = os.path.join(save_path, file_name + ".xlsx")

        # Bring Excel to foreground
        excel_window = wait_for_excel_open()
        if not excel_window:
            logging.error("Excel window not found for saving.")
            return False

        excel_window.activate()
        time.sleep(1)

        # Press F12 to open Save As
        pyautogui.press('f12')
        time.sleep(1)

        # Write full path
        pyautogui.hotkey('alt', 'd')  # Focus address bar
        time.sleep(0.2)
        pyautogui.write(save_path)
        pyautogui.press('enter')
        time.sleep(0.5)

        # Move to filename field
        pyautogui.press('tab', presses=6, interval=0.2)
        time.sleep(0.2)
        pyautogui.write(file_name, interval=0.05)
        time.sleep(0.5)

        # Press Save
        pyautogui.press('enter')
        logging.info(f"Excel file saved as {full_path}")
        return True

    except Exception as e:
        logging.error(f"Error while saving Excel file: {str(e)}")
        print(f"Error while saving Excel file: {str(e)}")
        return False

def save_and_close_excel(save_path, file_name, timeout=30):
    """Check if Excel file is saved, then terminate Excel process."""
    try:
        expected_file = os.path.join(save_path, file_name if file_name.lower().endswith('.xlsx') else file_name + '.xlsx')

        end_time = time.time() + timeout
        while time.time() < end_time:
            if os.path.exists(expected_file):
                logging.info(f"Confirmed file exists: {expected_file}")

                # Kill all EXCEL.EXE processes
                for proc in psutil.process_iter(['pid', 'name']):
                    if proc.info['name'] and 'excel' in proc.info['name'].lower():
                        logging.info(f"Terminating Excel process PID {proc.info['pid']}")
                        proc.terminate()

                time.sleep(2)  # Give Excel a moment to close
                return True

            time.sleep(1)

        logging.error(f"Excel file not found after {timeout} seconds: {expected_file}")
        print(f"Excel file not found after {timeout} seconds: {expected_file}")
        return False

    except Exception as e:
        logging.error(f"Error saving and closing Excel: {str(e)}")
        print(f"Error saving and closing Excel: {str(e)}")
        return False

def click_parameters(report_name, save_path, file_name):
    """Main function to open report, fill parameters, and handle export."""
    try:
        # Open report
        pyautogui.doubleClick(x=1001, y=96, duration=0)
        pyautogui.press('tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.press('backspace')
        pyautogui.write(report_name)
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.doubleClick(x=1071, y=194, duration=0)

        # First window - wait for report to load
        report_window = wait_for_window(report_name, 300)
        if not report_window:
            return False
        if not wait_for_window_to_stabilize(report_window):
            return False
        report_window.activate()
        time.sleep(0.5)

        # Click into parameters
        pyautogui.moveTo(report_window.left + 300, report_window.top + 80, duration=0)
        pyautogui.click()
        time.sleep(0.2)
        pyautogui.click()
        time.sleep(0.2)

        # Fill parameters
        pyautogui.press('backspace')
        pyautogui.write(start_date)
        pyautogui.press('tab')
        pyautogui.write(end_date)
        pyautogui.press('enter')
        logging.info("Parameters filled.")

        # Wait for Cube Viewer
        logging.info("Waiting for Cube Viewer window to appear...")
        report_window = wait_for_window("Cube Viewer", 1000)
        if not report_window:
            return False
        if not wait_for_window_to_stabilize(report_window):
            return False
        report_window.activate()
        time.sleep(0.5)

        # Now instead of clicking export, send Alt+F, E, E
        pyautogui.keyDown('alt')
        time.sleep(0.2)
        pyautogui.press('f')
        time.sleep(0.2)
        pyautogui.press('e')
        time.sleep(0.2)
        pyautogui.press('e')
        time.sleep(0.2)
        pyautogui.keyUp('alt')
        logging.info("Triggered export using Alt+F, E, E")

        # Wait for Excel
        if not wait_for_excel_open(60):
            logging.error("Excel did not open after export trigger.")
            return False

        # Save Excel file
        if not save_excel_file(save_path, file_name):
            logging.error("Failed to save Excel file.")
            return False

        # Now check the file is saved and close Excel
        if not save_and_close_excel(save_path, file_name):
            logging.error("Failed to properly close Excel.")
            return False

        # --- New part: After Excel closed, handle Cube Viewer and Fidelio ---
        # After closing Excel, reactivate Cube Viewer
        cube_viewer_window = wait_for_window("Cube Viewer", 10)
        if cube_viewer_window:
            cube_viewer_window.activate()
            time.sleep(0.5)
            pyautogui.press('esc')  # Close Cube Viewer
            logging.info("Closed Cube Viewer window with Esc.")
        else:
            logging.warning("Cube Viewer window not found after Excel close.")

        # Now activate the Fidelio window
        fidelio_window = wait_for_window("Fidelio", 10)
        if fidelio_window:
            fidelio_window.activate()
            logging.info("Fidelio window reactivated.")
        else:
            logging.warning("Fidelio window not found.")

        return True

    except Exception as e:
        logging.error(f"Error in click_parameters: {str(e)}")
        print(f"Error in click_parameters: {str(e)}")
        return False

# Example usage
if __name__ == "__main__":
    report_name = "occupancy history & forecast with budget"  # Report window title
    save_path = "D:\\Daily_Chart\\Chart_Resources\\S25\\HRG"   # Target Folder
    file_name = "DB H&F"  # Desired file name

    success = click_parameters(report_name, save_path, file_name)
    if success:
        logging.info("Automation completed successfully!")
    else:
        logging.error("Automation failed.")
