import appdirs
import logging
import logging.handlers
from pathlib import Path

APP_DIR = Path(appdirs.user_data_dir("wsl-usb-gui", False))
APP_DIR.mkdir(exist_ok=True)
LOG_FILE = APP_DIR / "log.txt"

formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler = logging.handlers.RotatingFileHandler(
    filename=LOG_FILE,  # Name of the log file
    maxBytes=1048576,   # Maximum file size (1 MB)
    backupCount=5       # Number of backup files to keep
)
file_handler.setFormatter(formatter)
stream_handler = logging.StreamHandler()
stream_handler.setFormatter(formatter)
log = logging.getLogger()
log.addHandler(file_handler)
log.addHandler(stream_handler)
log.setLevel(logging.INFO)  # Log INFO messages and above
