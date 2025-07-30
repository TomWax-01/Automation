import subprocess
import time
import sys
import logging
import psutil
import uiautomation as auto
import pyautogui
import pygetwindow as gw

from Main import username, password

# Configure logging
logging.basicConfig(filename="Report_Automation.log", level=logging.INFO,
                    format="%(asctime)s - %(message)s")

pyautogui.PAUSE = 0  # Speed up PyAutoGUI

main_window_gw = None  # Global main window object


def get_window_by_title(title):
    try:
        windows = gw.getWindowsWithTitle(title)
        for window in windows:
            if title in window.title:
                return window
    except Exception as e:
        logging.error(f"Error finding window: {str(e)}")
    return None


def kill_existing_fidelio_processes():
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and proc.info['name'].lower() == 'fideliov8.exe':
            proc.terminate()
            proc.wait()
            logging.info("Existing fideliov8.exe process terminated.")
            print("Existing fideliov8.exe process terminated.")


def list_open_windows():
    root = auto.GetRootControl()
    print("Listing all open windows:")
    for child in root.GetChildren():
        if child.Name:
            print(child.Name)


def get_top_windows():
    root = auto.GetRootControl()
    return [control for control in root.GetChildren() if control.ControlTypeName == 'WindowControl']


def handle_error_dialogs(main_window, timeout=5):
    start_time = time.time()
    error_count = 0

    while time.time() - start_time < timeout:
        try:
            error_windows = get_top_windows()
            for window in error_windows:
                if window.Name and "Oracle Hospitality Suite8" in window.Name and window != main_window:
                    try:
                        ok_button = window.ButtonControl(Name="OK")
                        if ok_button.Exists(0, 0):
                            print(f"Found and dismissing error dialog: {window.Name}")
                            logging.info(f"Dismissing error dialog: {window.Name}")
                            ok_button.Click()
                            error_count += 1
                            time.sleep(0.2)
                            continue
                        window.SetFocus()
                        window.SendKeys("{ENTER}")
                        print(f"Dismissed error dialog with ENTER: {window.Name}")
                        error_count += 1
                        time.sleep(0.2)
                    except Exception as e:
                        print(f"Error handling dialog: {str(e)}")
        except Exception as e:
            print(f"Error in dialog detection: {str(e)}")

        # Break early if no more error windows
        if len([w for w in get_top_windows() if w.Name and "Oracle Hospitality Suite8" in w.Name]) <= 1:
            break

        time.sleep(0.2)

    if error_count > 0:
        print(f"Handled {error_count} error dialog(s)")
        logging.info(f"Dismissed {error_count} error dialog(s)")
    else:
        print("No error dialogs detected")
        logging.info("No error dialogs detected")

    return error_count


def launch_fidelio_with_config(executable_path, config_file_path, username, password):
    global main_window_gw

    try:
        kill_existing_fidelio_processes()
        time.sleep(0.5)

        subprocess.Popen([executable_path, config_file_path])
        logging.info("Fidelio application launched.")
        print("Fidelio application launched.")

        # Wait for login window
        login_window = auto.WindowControl(Name="Oracle Hospitality Suite8 Login")
        if not login_window.Exists(maxSearchSeconds=30):
            print("Login window not found within timeout.")
            logging.error("Login window not found within timeout.")
            sys.exit(1)

        print("Login window detected.")
        logging.info("Login window detected.")

        login_window.SetFocus()
        time.sleep(0.3)

        pyautogui.write(username, interval=0.01)
        pyautogui.press('tab')
        pyautogui.write(password, interval=0.01)
        pyautogui.press('enter')
        print("Credentials entered.")
        # Wait a few seconds for login to process
        time.sleep(5)

        # Check if login window still exists
        if login_window.Exists():
            print("Login window still exists after credentials input - possible login failure.")
            logging.error("Login failed: Login window still open after credentials entered.")
            sys.exit(1)
        else:
            print("Login successful. Continuing...")
            logging.info("Login successful.")

        # Wait for the main window
        main_window = auto.WindowControl(Name="Oracle Hospitality Suite8", searchDepth=1, SubName=True)
        if not main_window.Exists(maxSearchSeconds=10):
            print("Main window not found, trying backup method...")
            logging.warning("Main window not found, trying backup.")
            main_window_gw = get_window_by_title("Oracle Hospitality Suite8")
            if not main_window_gw:
                print("Main window still not found.")
                logging.error("Main window still not found.")
                sys.exit(1)
        else:
            main_window.SetFocus()
            main_window.Maximize()
            print("Main window focused.")
            windows = gw.getWindowsWithTitle("Oracle Hospitality Suite8")
            if windows:
                main_window_gw = windows[0]

        if main_window_gw:
            main_window_gw.activate()
            pyautogui.click(main_window_gw.left + 1, main_window_gw.top + 50)
            print("Main window activated.")
            time.sleep(0.5)

        # Handle any error dialogs
        if main_window.Exists():
            ui_main_window = main_window
        elif main_window_gw:
            try:
                ui_main_window = auto.ControlFromHandle(main_window_gw._hWnd)
            except Exception as e:
                logging.error(f"Handle conversion error: {str(e)}")
                ui_main_window = None
        else:
            ui_main_window = None

        if ui_main_window:
            handle_error_dialogs(ui_main_window, timeout=2)

        # Send ALT + I + R
        pyautogui.hotkey('altleft', 'i', 'r')
        print("Sent ALT + I + R")

        # Click on specific offset
        if main_window_gw:
            pyautogui.moveTo(main_window_gw.left + (1228 - main_window_gw.left),
                             main_window_gw.top + (97 - main_window_gw.top))
            pyautogui.click()
            print("Clicked at target offset.")
        else:
            print("Window not found for offset click.")
            logging.error("Window not found for offset click.")
            sys.exit(1)

    except Exception as e:
        logging.error(f"Launch error: {str(e)}")
        print(f"Launch error: {str(e)}")
        sys.exit(1)


def launch_fidelio():
    executable_path = r"C:\FIDELIO\Programs\fideliov8.exe"
    config_file_path = r"C:\FIDELIO\v8live.ini"

    launch_fidelio_with_config(executable_path, config_file_path, username, password)


if __name__ == "__main__":
    launch_fidelio()
