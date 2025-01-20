# Windows service for auto-rsa (discord)
# pip install pywin32
# python autorsa_service.py install 

import win32serviceutil
import win32service
import win32event
import logging
import os
import subprocess

class AutoRSAService(win32serviceutil.ServiceFramework):
    _svc_name_ = "AutoRSAService"
    _svc_display_name_ = "Auto RSA Service"
    _svc_description_ = "Runs the autoRSA.py script with 'discord' argument as a service."

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.running = True
        # Set up logging
        self.setup_logging()

    def setup_logging(self):
        log_folder = r"C:\Users\Admin\Logs" # Change it with your path
        os.makedirs(log_folder, exist_ok=True)  # Ensure the log folder exists
        log_file = os.path.join(log_folder, "autoRSA_service.log")
        logging.basicConfig(
            filename=log_file,
            level=logging.DEBUG,
            format="%(asctime)s - %(levelname)s - %(message)s"
        )
        logging.info("Service initialized.")

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.running = False
        logging.info("Service stopped.")

    def SvcDoRun(self):
        logging.info("Service started.")
        python_path = r"C:\Users\Admin\AppData\Local\Programs\Python\Python311\python.exe" # Change it with your path
        script_path = r"C:\Users\Admin\auto-rsa\autoRSA.py" # Change it with your path
        args = "discord"
        cwd_path = r"C:\Users\Admin\auto-rsa"  # Set the working directory to the repository path

        try:
            # Log environment details
            env = os.environ.copy()
            logging.info(f"Environment: {env}")

            # Run the script
            logging.info(f"Running script: {python_path} {script_path} {args} with cwd={cwd_path}")
            process = subprocess.Popen(
                [python_path, script_path, args],
                cwd=cwd_path,
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            stdout, stderr = process.communicate()

            # Log script output and errors
            if stdout:
                logging.info(f"Script output: {stdout}")
            if stderr:
                logging.error(f"Script error: {stderr}")

        except Exception as e:
            logging.error(f"Error running the script: {e}")

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(AutoRSAService)
