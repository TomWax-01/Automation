import pyautogui
import time
import pygetwindow as gw
import logging
import os

from Main import start_date, end_date


log_file = os.path.join(os.path.dirname(__file__), "Report_Automation.log")
logging.basicConfig(
    filename=log_file,
    filemode='a',  # Append to the log file. Use 'w' to overwrite each time
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

def close_report_window(report_name, max_wait_time=30):
    """Close report window by sending ESC repeatedly."""
    try:
        report_window = wait_for_window(report_name, max_wait_time)
        if not report_window:
            logging.error(f"Could not find report window to close: {report_name}")
            print(f"Could not find report window to close: {report_name}")
            return False

        if not wait_for_window_to_stabilize(report_window, max_wait_time):
            logging.error(f"Report window not responsive for closing: {report_name}")
            print(f"Report window not responsive for closing: {report_name}")
            return False

        report_window.activate()
        time.sleep(1)

        max_esc_presses = 10
        esc_count = 0

        while esc_count < max_esc_presses:
            if not get_window_by_title(report_name):
                logging.info(f"Report window successfully closed after {esc_count} ESC presses")
                print(f"Report window successfully closed after {esc_count} ESC presses")
                return True

            pyautogui.press('escape')
            logging.info(f"Pressed ESC key ({esc_count + 1}/{max_esc_presses})")
            print(f"Pressed ESC key ({esc_count + 1}/{max_esc_presses})")
            esc_count += 1
            time.sleep(3)

        if get_window_by_title(report_name):
            logging.warning(f"Report window still exists after {max_esc_presses} ESC presses")
            print(f"Warning: Report window still exists after {max_esc_presses} ESC presses")
            return False
        else:
            logging.info("Report window closed successfully")
            print("Report window closed successfully")
            return True

    except Exception as e:
        logging.error(f"Error closing report window: {str(e)}")
        print(f"Error closing report window: {str(e)}")
        return False

def handle_export_report(save_path, file_name, report_name=None):
    """Handle Export Report Save Dialog"""
    try:
        export_window = wait_for_window('Export Report', 30)
        if not export_window:
            return False
        if not wait_for_window_to_stabilize(export_window, 10):
            return False

        export_window.activate()
        time.sleep(0.5)

        # Click address bar (shortcut)
        pyautogui.hotkey('alt', 'd')
        time.sleep(0.2)
        pyautogui.write(save_path)
        pyautogui.press('enter')
        logging.info(f"Changed save location to {save_path}.")
        time.sleep(1)

        # File name field
        pyautogui.press('tab', presses=6, interval=0.1)
        time.sleep(0.2)
        pyautogui.write(file_name, interval=0.05)
        logging.info(f"Set file name to {file_name}.")
        time.sleep(0.5)

        # Save as type dropdown
        pyautogui.press('tab')
        time.sleep(0.2)
        pyautogui.press('down', presses=4, interval=0.1)
        pyautogui.press('enter')
        time.sleep(0.5)

        # Move to Save button
        pyautogui.press('tab', presses=2, interval=0.1)
        pyautogui.press('enter')
        pyautogui.hotkey('alt', 'y')
        time.sleep(0.5)
        logging.info("Clicked Save.")

        # Close the report window after saving
        if report_name:
            close_report_window(report_name)

        time.sleep(1)
        return True

    except Exception as e:
        logging.error(f"Error in handle_export_report: {str(e)}")
        print(f"Error in handle_export_report: {str(e)}")
        return False

def click_parameters(report_name, save_path, file_name):
    """Main function to open report, fill parameters, and handle export."""
    try:
        # Open report
        pyautogui.doubleClick(x=1001, y=96, duration=0)
        pyautogui.press('tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.press('backspace')
        pyautogui.write(report_name, interval=0)
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.doubleClick(x=1071, y=194, duration=0)

        # First window - increased timeout to 100
        report_window = wait_for_window(report_name, 100)
        if not report_window:
            return False
        if not wait_for_window_to_stabilize(report_window):
            return False
        report_window.activate()
        time.sleep(0.3)

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

        # Wait for window reload and ensure it's responsive
        logging.info("Waiting for report window to reload...")
        report_window = wait_for_window(report_name, 100)
        if not report_window:
            return False

        # Multiple attempts to check window responsiveness and trigger export
        max_attempts = 5
        for attempt in range(max_attempts):
            if not wait_for_window_to_stabilize(report_window):
                logging.info(f"Window not stable yet, attempt {attempt + 1}/{max_attempts}")
                time.sleep(2)

                # Verify window still exists
                if not get_window_by_title(report_window.title):
                    logging.error("Window no longer exists")
                    return False
                continue

            try:
                report_window.activate()
                time.sleep(0.5)

                # Multiple checks for window state
                current_window = get_window_by_title(report_window.title)
                if (current_window and
                    current_window.isActive and
                    current_window.visible):

                    logging.info("Window is stable and responsive")
                    break
                else:
                    logging.warning("Window failed state checks")
                    time.sleep(1)

            except Exception as e:
                logging.warning(f"Window not yet responsive, attempt {attempt + 1}/{max_attempts}: {str(e)}")
                time.sleep(1)
        else:
            logging.error("Window failed to become responsive")
            return False

        # Click export trigger repeatedly until export window appears
        max_export_attempts = 21  # Try for 21 seconds total
        export_click_x = report_window.left + 11
        export_click_y = report_window.top + 35

        for attempt in range(max_export_attempts):
            # Click the export trigger
            pyautogui.moveTo(export_click_x, export_click_y, duration=0)
            pyautogui.click()
            logging.info(f"Export trigger click attempt {attempt + 1}/{max_export_attempts}")

            # Check if export window appeared
            export_window = wait_for_window('Export Report', 1)
            if export_window:
                logging.info("Export window appeared successfully")
                break

            time.sleep(1)
        else:
            logging.error("Export window failed to appear after multiple attempts")
            return False

        # Handle Export Report window
        if not handle_export_report(save_path, file_name, report_name):
            return False

        return True

    except Exception as e:
        logging.error(f"Error in click_parameters: {str(e)}")
        print(f"Error in click_parameters: {str(e)}")
        return False

# Example usage
if __name__ == "__main__":
    report_name = "Forecast - Occupancy/Availability per Room Type"  # Report window title
    save_path = "D:\\Daily_Chart\\Chart_Resources\\S25\\HRG"  # Target Folder
    file_name = "DB RTA"  # Desired file name

    success = click_parameters(report_name, save_path, file_name)
    if success:
        logging.info("Automation completed successfully!")
    else:
        logging.error("Automation failed.")
