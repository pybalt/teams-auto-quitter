import os
import platform
import subprocess
import time
import datetime
import threading
import sys
import signal

# Tuple of Microsoft Teams related services and processes
TEAMS_PROCESSES = (
    "Teams.exe",
    "Microsoft.Teams.exe",
    "Teams",
    "Microsoft Teams",
    "Teams Helper",
    "Teams Updater",
    "Teams Background Service",
    "Microsoft Teams Update Service",
    "Teams Meeting Addin",
    "Teams.AudioService.exe",
    "Teams.VideoService.exe",
    "msedgewebview2.exe",
    "crashpad_handler.exe",
    "RuntimeBroker.exe"
)

def check_teams_processes():
    """Check if any Microsoft Teams related processes are running."""
    print("Checking for Microsoft Teams processes...")
    
    system = platform.system()
    running_teams_processes = []
    
    if system == "Windows":
        try:
            # Using tasklist command on Windows
            output = subprocess.check_output("tasklist", shell=True).decode('utf-8', errors='ignore')
            for process in TEAMS_PROCESSES:
                if process.lower() in output.lower():
                    running_teams_processes.append(process)
        except Exception as e:
            print(f"Error checking processes: {e}")
    
    elif system == "Linux" or system == "Darwin":  # Linux or macOS
        try:
            # Using ps command on Unix-like systems
            cmd = "ps aux" if system == "Linux" else "ps -ax"
            output = subprocess.check_output(cmd, shell=True).decode('utf-8', errors='ignore')
            for process in TEAMS_PROCESSES:
                if process.lower() in output.lower():
                    running_teams_processes.append(process)
        except Exception as e:
            print(f"Error checking processes: {e}")
    
    else:
        print(f"Unsupported operating system: {system}")
    
    if running_teams_processes:
        print("Found Teams processes running:")
        for process in running_teams_processes:
            print(f"- {process}")
    else:
        print("No Microsoft Teams processes found running.")
    
    return running_teams_processes

def terminate_teams_processes():
    """Terminate all Microsoft Teams related processes."""
    print("Attempting to terminate Microsoft Teams processes...")
    
    system = platform.system()
    running_processes = check_teams_processes()
    
    if not running_processes:
        print("No Teams processes to terminate.")
        return
    
    if system == "Windows":
        try:
            for process in running_processes:
                print(f"Terminating {process}...")
                subprocess.run(f"taskkill /F /IM \"{process}\"", shell=True)
            print("Teams processes termination attempted.")
        except Exception as e:
            print(f"Error terminating processes: {e}")
    
    elif system == "Linux":
        try:
            for process in running_processes:
                print(f"Terminating {process}...")
                subprocess.run(f"pkill -f \"{process}\"", shell=True)
            print("Teams processes termination attempted.")
        except Exception as e:
            print(f"Error terminating processes: {e}")
    
    elif system == "Darwin":  # macOS
        try:
            for process in running_processes:
                print(f"Terminating {process}...")
                subprocess.run(f"killall \"{process}\"", shell=True)
            print("Teams processes termination attempted.")
        except Exception as e:
            print(f"Error terminating processes: {e}")
    
    else:
        print(f"Unsupported operating system: {system}")
    
    # Verify termination
    remaining = check_teams_processes()
    if not remaining:
        print("All Teams processes successfully terminated.")
    else:
        print(f"Some Teams processes could not be terminated: {remaining}")

def list_running_processes():
    """List all running processes/programs on the system."""
    print("Running processes/programs:")
    
    system = platform.system()
    
    if system == "Windows":
        try:
            # Using tasklist command on Windows
            output = subprocess.check_output("tasklist", shell=True).decode('utf-8', errors='ignore')
            print(output)
        except Exception as e:
            print(f"Error getting process list: {e}")
    
    elif system == "Linux" or system == "Darwin":  # Linux or macOS
        try:
            # Using ps command on Unix-like systems
            cmd = "ps aux" if system == "Linux" else "ps -ax"
            output = subprocess.check_output(cmd, shell=True).decode('utf-8', errors='ignore')
            print(output)
        except Exception as e:
            print(f"Error getting process list: {e}")
    
    else:
        print(f"Unsupported operating system: {system}")

def wait_until_time(target_hour, target_minute=0):
    """Wait until specified time is reached.
    
    Args:
        target_hour (int): Hour to wait for (24-hour format)
        target_minute (int, optional): Minute to wait for. Defaults to 0.
    """
    while True:
        now = datetime.datetime.now()
        if now.hour > target_hour or (now.hour == target_hour and now.minute >= target_minute):
            print(f"It's now {now.hour}:{now.minute}, executing Teams termination.")
            break
        else:
            print(f"Current time: {now.hour}:{now.minute}. Waiting until {target_hour}:{target_minute}...")
            time.sleep(60)  # Check every minute

def run_in_background(target_hour, target_minute):
    """Run the Teams termination process in the background."""
    def background_task():
        print(f"Background task started. Will terminate Teams processes at {target_hour}:{target_minute}.")
        wait_until_time(target_hour, target_minute)
        terminate_teams_processes()
        print("Background task completed.")
    
    # Create and start the background thread
    background_thread = threading.Thread(target=background_task, daemon=True)
    background_thread.start()
    return background_thread

def handle_exit(signum, frame):
    """Handle exit signals gracefully."""
    print("\nExiting Teams terminator. Background task will continue running.")
    sys.exit(0)

if __name__ == "__main__":
    # Prompt user for termination time
    try:
        target_hour = int(input("Enter the hour to terminate Teams (0-23): "))
        while target_hour < 0 or target_hour > 23:
            target_hour = int(input("Invalid hour. Please enter a value between 0-23: "))
            
        target_minute = int(input("Enter the minute to terminate Teams (0-59): "))
        while target_minute < 0 or target_minute > 59:
            target_minute = int(input("Invalid minute. Please enter a value between 0-59: "))
    except ValueError:
        print("Invalid input. Using default values: 13:00")
        target_hour = 13
        target_minute = 0
    
    # Set up signal handlers for graceful exit
    signal.signal(signal.SIGINT, handle_exit)
    signal.signal(signal.SIGTERM, handle_exit)
    
    print(f"Program started. Will terminate Teams processes at {target_hour}:{target_minute}.")
    print("This program will run in the background. You can continue using your PC.")
    
    # Run the termination task in the background
    background_thread = run_in_background(target_hour, target_minute)
    
    # If on Windows, hide the console window
    if platform.system() == "Windows":
        try:
            # This will minimize the console window
            subprocess.run("cmd /c start /min cmd", shell=True)
        except Exception as e:
            print(f"Error minimizing window: {e}")
    
    # Keep the main thread alive but not consuming CPU
    try:
        while background_thread.is_alive():
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nProgram interrupted. Background task will continue running.")
