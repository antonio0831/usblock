import time
import ctypes
import win32com.client
import yaml
import os
import logging

from win10toast import ToastNotifier

TARGET_DEVICE_ID = ""
CHECK_INTERVAL = 1
MAX_LOCK_COUNT = 0
LOG_INTERVAL = 300
CURRENT_VERSION = 0.1
# Function to create a default config.yml if it doesn't exist
def create_default_config(config_file='config.yml'):
    default_config = {
        'usb_id': '',  # the id for the usb, under device manager
        'watchdog_interval': 1,  # in seconds, to check if USB is connected or not
        'max_lockout': 0,  # the max lockout times, 0 means unlimited
        'log_interval': 300  # in seconds, 0 means no log
    }
    with open(config_file, 'w') as file:
        yaml.dump(default_config, file)
    print(f"Default configuration file created at {config_file}")


# Function to read configuration from config.yml
def read_config(config_file='config.yml'):
    if not os.path.exists(config_file):
        create_default_config(config_file)
    with open(config_file, 'r') as file:
        config = yaml.safe_load(file)
    return config

def send_notification(title, message):
    toaster = ToastNotifier()
    toaster.show_toast(title, message, duration=5)  # duration is in seconds
def load_config():
    global TARGET_DEVICE_ID, CHECK_INTERVAL, MAX_LOCK_COUNT, LOG_INTERVAL
    config = read_config()
    TARGET_DEVICE_ID = config['usb_id'].upper()
    CHECK_INTERVAL = config['watchdog_interval']
    MAX_LOCK_COUNT = config['max_lockout']
    LOG_INTERVAL = config['log_interval']


def lock_computer():
    ctypes.windll.user32.LockWorkStation()

def setup_logging():
    # Retrieve the path using the APPDATA environment variable
    appdata_path = os.getenv('APPDATA')

    # Define a subdirectory and file name within the AppData folder
    myapp_path = os.path.join(appdata_path, 'USBWatch')
    os.makedirs(myapp_path, exist_ok=True)

    # Full path for the file
    log_file = os.path.join(myapp_path, 'system.log')
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    logging.info('-------------------------------------------------')

def is_target_device_connected(target_device_id):
    target_device_id = target_device_id.upper()
    wmi = win32com.client.GetObject("winmgmts:")
    devices = wmi.ExecQuery("SELECT * FROM Win32_PnPEntity")
    for device in devices:
        device_id = device.DeviceID.upper()  # Ensure case-insensitive comparison
        if target_device_id == device_id:
            return True
    return False


def main():
    load_config()
    lock_count = 0
    was_connected = False  # Track if the device was connected at least once
    setup_logging()
    last_log_time = time.time()
    send_notification("USB Watchdog", "System Initiated")
    logging.info(f"Started usb watchdog under {os.getlogin()}")
    logging.info(f"Version: {CURRENT_VERSION}")
    logging.info(f"Config Status:")
    logging.info(f"USB ID: {TARGET_DEVICE_ID}")
    logging.info(f"CHECK Interval: {CHECK_INTERVAL}")
    logging.info(f"Log Interval: {LOG_INTERVAL}")
    logging.info(f"Max Lockout: {MAX_LOCK_COUNT}")
    while MAX_LOCK_COUNT == 0 or lock_count < MAX_LOCK_COUNT:
        currently_connected = is_target_device_connected(TARGET_DEVICE_ID)
        if currently_connected:
            was_connected = True
        elif was_connected:
            logging.info("Device disconnected, locking computer")
            lock_computer()
            lock_count += 1
            was_connected = False  # Reset tracking as device is now removed
        if LOG_INTERVAL > 0 and (time.time() - last_log_time >= LOG_INTERVAL):
            logging.info(f"System is operational. Attached: {currently_connected}; Lock Count: {lock_count}")
            last_log_time = time.time()

        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    main()
