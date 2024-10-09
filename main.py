import subprocess
import re
import logging

logging.basicConfig(
    format="%(message)s",
    level=logging.INFO,
)


class Colors:
    HEADER = "\033[95m"
    OKBLUE = "\033[94m"
    OKCYAN = "\033[96m"
    OKGREEN = "\033[92m"
    WARNING = "\033[93m"
    FAIL = "\033[91m"
    ENDC = "\033[0m"
    BOLD = "\033[1m"
    UNDERLINE = "\033[4m"


def clear_screen():
    subprocess.run(["cls"], shell=True)


def log_colored(message, level, color):
    colored_message = f"{color}{message}{Colors.ENDC}"
    if level == "info":
        logging.info(colored_message)
    elif level == "warning":
        logging.warning(colored_message)
    elif level == "error":
        logging.error(colored_message)


def get_time_input(prompt):
    time_input = input(prompt)
    return time_input.zfill(2)


def validate_time_format(time):
    return re.match(r"^(0[0-9]|1[0-9]|2[0-3]):[0-5][0-9]$", time)


def time_format():
    time_hour = get_time_input("Enter time hour (00-23): ")
    time_minute = get_time_input("Enter time minute (00-59): ")
    time = f"{time_hour}:{time_minute}"

    if not validate_time_format(time):
        return False

    return time


def vallidate_script_path(script_path):
    if not script_path.endswith(".ps1"):
        log_colored(
            "Invalid script path. PowerShell script should have a .ps1 extension.",
            "error",
            Colors.FAIL,
        )
        return False

    if " " in script_path:
        script_path = f'\\"{script_path}\\"'

    return script_path


def schedule_task(folder_path):
    clear_screen()
    task_name = input("Enter the name for the task: ")
    script_path = input("Enter the full path of the PowerShell script: ")
    if not vallidate_script_path(script_path):
        return
    frequency = input("Enter the frequency (daily, weekly, monthly): ").lower()
    time = time_format()

    if not time:
        log_colored(
            "Invalid time format. Please enter a valid time.", "error", Colors.FAIL
        )
        return

    command = [
        "schtasks",
        "/create",
        "/tn",
        folder_path + task_name,
        "/tr",
        f'powershell.exe -File "{script_path}"',
        "/sc",
        frequency,
        "/st",
        time,
    ]

    try:
        subprocess.run(command, check=True)
        clear_screen()
        log_colored(
            f"Task '{task_name}' scheduled successfully for {time} - {frequency}.",
            "info",
            Colors.OKGREEN,
        )
    except subprocess.CalledProcessError as e:
        clear_screen()
        log_colored(f"Failed to schedule task: {e}", "error", Colors.FAIL)


def list_existing_tasks(folder_path):
    list_command = [
        "schtasks",
        "/query",
        "/fo",
        "list",
        "/v",
        "/tn",
        f"{folder_path}",
    ]
    existing_tasks = subprocess.run(
        list_command, check=True, stdout=subprocess.PIPE, text=True
    )
    return existing_tasks.stdout.strip()


def delete_existing_tasks(folder_path):
    clear_screen()
    log_colored("\nExisting Scheduled Tasks:", "info", Colors.HEADER)
    tasks = []
    try:
        output = list_existing_tasks(folder_path)
        if not output:
            log_colored("No tasks found.", "warning", Colors.WARNING)
            return

        output = output.split("\n")

        task_index = 0
        for line in output:
            if line.startswith("TaskName"):
                task_name = line.split()[-1].strip()
                task_index += 1
                log_colored(
                    f"\n{task_index}) Task Name: {task_name}", "info", Colors.OKBLUE
                )
                tasks.append(task_name)

        task_index = input("\nEnter the task number to delete: ")
        if not task_index.isdigit() or int(task_index) > len(tasks):
            log_colored(
                "Invalid task number. Please enter a valid task number.",
                "error",
                Colors.FAIL,
            )
            return

        task_name = tasks[int(task_index) - 1]
        delete_command = ["schtasks", "/delete", "/tn", task_name, "/f"]
        clear_screen()
        subprocess.run(delete_command, check=True)
        log_colored(f"Task '{task_name}' deleted successfully.", "info", Colors.OKGREEN)
    except subprocess.CalledProcessError as e:
        log_colored(f"Failed to delete task: {e}", "error", Colors.FAIL)


def update_folder_path():
    folder = input("\nEnter the folder name from Task Scheduler Application: ")

    if not folder:
        log_colored("Please enter a valid folder name.", "error", Colors.FAIL)
        return

    if not folder.endswith("\\"):
        folder += "\\"

    check_folder_exists = subprocess.run(
        ["schtasks", "/query", "/tn", folder], stderr=subprocess.PIPE
    )
    if check_folder_exists.returncode == 0:
        clear_screen()
        log_colored(f"Folder path updated to {folder}.", "info", Colors.OKGREEN)
        with open("config.txt", "w") as file:
            file.write(folder)
        return 0

    elif check_folder_exists.returncode == 1:
        clear_screen()
        log_colored(
            f"Folder name '{folder}' does not exist. Please enter a valid folder name from Task Scheduler Application.",
            "error",
            Colors.FAIL,
        )
        return 1

    clear_screen()
    log_colored(f"Folder path updated to {folder}.", "info", Colors.OKGREEN)


def read_folder_path():
    try:
        with open("config.txt", "r") as file:
            return file.read().strip()
    except FileNotFoundError:
        log_colored(
            "Initial setup required. Please enter the task scheduler folder name.",
            "warning",
            Colors.WARNING,
        )
        while True:
            if update_folder_path() == 0:
                return read_folder_path()


def main():
    while True:
        folder_path = read_folder_path()

        print("\nMain Menu:")
        print("1. Schedule a New Task")
        print("2. Delete Existing Task")
        print("3. Change Folder")
        print("4. Exit")

        choice = input("Select an option (1-4): ")

        if choice == "1":
            schedule_task(folder_path)
        elif choice == "2":
            delete_existing_tasks(folder_path)
        elif choice == "3":
            update_folder_path()
        elif choice == "4":
            print("\nExiting the program.")
            break
        else:
            log_colored(
                "Invalid choice. Please select a valid option.", "error", Colors.FAIL
            )


if __name__ == "__main__":
    main()
