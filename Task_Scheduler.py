import time
import smtplib
import win32com.client
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import ctypes
from datetime import datetime, timezone
from datetime import timedelta
import json

# Configuration for SMTP
SMTP_SERVER = 'smtp.office365.com'
SMTP_PORT = 587
USERNAME = 'noreply@fi911.com'
PASSWORD = 'Lop72698'
RECEIVER_EMAIL = 'y.raj@911fintech.com'
ENVIRONMENT = "Sandbox"

aree bhai

IST_OFFSET = timedelta(hours=5, minutes=30)

# Define expected runtime for each task (in seconds)
def load_expected_runtime():
    with open('expected_runtime.json', 'r') as f:
        return json.load(f)
EXPECTED_RUNTIME = load_expected_runtime()  # Load runtimes from JSON file

STATUS_MAPPING = {
    0: "Unknown",
    1: "Disabled",
    2: "Queued",
    3: "Ready",
    4: "Running"   
}

def send_email(subject, body):
    msg = MIMEMultipart()
    msg['From'] = USERNAME
    msg['To'] = RECEIVER_EMAIL
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'html'))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(USERNAME, PASSWORD)
        server.sendmail(msg['From'], msg['To'], msg.as_string())


def get_error_message(error_code):
    FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x00000100
    FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000
    FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200

    # Allocate a buffer for the message
    lpMsgBuf = ctypes.c_void_p()

    # Attempt to format the message
    length = ctypes.windll.kernel32.FormatMessageW(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        None,
        error_code,
        0,
        ctypes.byref(lpMsgBuf),
        0,
        None
    )

    if length == 0:
        return "Unknown error occurred."

    # Extract the message and free the buffer
    error_message = ctypes.wstring_at(lpMsgBuf)
    ctypes.windll.kernel32.LocalFree(lpMsgBuf)

    return error_message.strip()  # Return the message as it is

def check_tasks():
    scheduler = win32com.client.Dispatch("Schedule.Service")
    scheduler.Connect()
    rootFolder = scheduler.GetFolder("\\")
    tasks = rootFolder.GetTasks(0)
    tasks_with_issues = []

    for task in tasks:
        task_name = task.Name
        task_info = task.State
        task_last_task_result = task.LastTaskResult
        task_last_run_time = task.LastRunTime
        
        print(f"Checking task: {task_name} (State: {task_info})")

        if task_name in EXPECTED_RUNTIME:
            expected_run_time = EXPECTED_RUNTIME[task_name]
            current_time_utc = datetime.now(timezone.utc)
            current_time_ist = current_time_utc + IST_OFFSET

            if task_info == 4:  # Task is running
                if task_last_run_time:
                    error_message = get_error_message(task_last_task_result)
                    if task_last_run_time.tzinfo is None:
                        task_last_run_time = task_last_run_time.replace(tzinfo=timezone.utc)

                    run_time = (current_time_ist - task_last_run_time).total_seconds()
                    if run_time > expected_run_time:
                        tasks_with_issues.append({
                            "task_name": task_name,
                            "expected_runtime": expected_run_time // 60,
                            "actual_runtime": int(run_time // 60),
                            "error_message": error_message,
                            "last_run_time": task_last_run_time,
                            "Status": "Running",
                            "environment": ENVIRONMENT
                        })

            elif task_info == 3:  # Task is ready (not running)
                if task_last_run_time:
                    if task_last_run_time.tzinfo is None:
                        task_last_run_time = task_last_run_time.replace(tzinfo=timezone.utc)

                    run_time = (current_time_ist - task_last_run_time).total_seconds()
                    if run_time < expected_run_time and task_last_task_result != 0:  # Not a successful completion
                        tasks_with_issues.append({
                            "task_name": task_name,
                            "expected_runtime": expected_run_time // 60,
                            "actual_runtime": int(run_time // 60),
                            "error_message": get_error_message(task_last_task_result),
                            "last_run_time": task_last_run_time,
                            "Status": "Failed",
                            "environment": ENVIRONMENT
                        })

    return tasks_with_issues

def gather_task_statuses():
    scheduler = win32com.client.Dispatch("Schedule.Service")
    scheduler.Connect()
    rootFolder = scheduler.GetFolder("\\")
    tasks = rootFolder.GetTasks(0)
    
    status_report = []

    for task in tasks:
        task_name = task.Name
        if task_name in EXPECTED_RUNTIME:
            task_info = task.State
            last_run_time = task.LastRunTime
            next_run_time = task.NextRunTime
            expected_runtime = EXPECTED_RUNTIME.get(task_name, "N/A")

            status_report.append({
                "task_name": task_name,
                "status": STATUS_MAPPING.get(task_info, "Unknown"),
                "last_run_time": last_run_time,
                "next_run_time": next_run_time,
                "expected_runtime": expected_runtime,
                "environment": ENVIRONMENT
            })

    return status_report

if __name__ == "__main__":
    all_tasks_with_issues = []
    start_time = time.time()
    alerted_tasks = {}
    status_email_time = time.time()

    while True:
        tasks_with_issues = check_tasks()

        for task in tasks_with_issues:
            task_name = task['task_name']
            if task_name not in alerted_tasks:
                all_tasks_with_issues.append(task)
                alerted_tasks[task_name] = True

        if time.time() - start_time >= 150: # Error Mail trigger time
            if all_tasks_with_issues:
                body = """
                <html>
                    <body>
                        <p>Hi,</p>
                        <p style="color:red;"><b>Warning:</b> The following tasks encountered issues:</p>
                        <table border="1">
                            <tr>
                                <th>Task Name</th>
                                <th>Status</th>
                                <th>Environment</th>
                                <th>Expected Runtime (min)</th>
                                <th>Actual Runtime (min)</th>
                                <th>Last Run Time</th>
                                <th>Error Message</th>
                            </tr>
                """
                for task in all_tasks_with_issues:
                    body += f"""
                    <tr>
                        <td>{task['task_name']}</td>
                        <td>{task['Status']}</td>
                        <td>{task['environment']}</td>
                        <td>{task['expected_runtime']}</td>
                        <td>{task['actual_runtime']}</td>
                        <td>{task['last_run_time'].strftime('%Y-%m-%d %H:%M:%S')}</td>
                        <td>{task['error_message']}</td>
                    </tr>
                    """
                body += """
                        </table>
                        <p>Regards-Fi911 Support</p>  <!-- Closing Remarks -->
                        <p style="color:red;"><b>Note:</b> This is an automated message, please do not reply.</p>
                    </body>
                </html>
                """
                send_email("Task Status Alert: Issues Detected", body)
                print("Alert email sent for accumulated tasks with issues.")

            all_tasks_with_issues.clear()
            start_time = time.time()
            alerted_tasks.clear()

        if time.time() - status_email_time >= 60: # Status Mail trigger time
            task_statuses = gather_task_statuses()
            status_body = """
            <html>
                <body>
                    <p>Hi,</p>
                    <p>Here are the current task statuses:</p>
                    <table border="1">
                        <tr>
                            <th>Task Name</th>
                            <th>Status</th>
                            <th>Environment</th>
                            <th>Last Run Time</th>
                            <th>Next Run Time</th>
                        </tr>
            """
            for task in task_statuses:
                last_run_time = task['last_run_time'].strftime('%Y-%m-%d %H:%M:%S') if task['last_run_time'] else "N/A"
                next_run_time = task['next_run_time'].strftime('%Y-%m-%d %H:%M:%S') if task['next_run_time'] else "N/A"
                expected_runtime = task['expected_runtime'] // 60 if isinstance(task['expected_runtime'], int) else "N/A"
                status_body += f"""
                <tr>
                    <td>{task['task_name']}</td>
                    <td>{task['status']}</td>
                    <td>{task['environment']}</td>
                    <td>{last_run_time}</td>
                    <td>{next_run_time}</td>
                </tr>
                """
            status_body += """
                    </table>
                    <p>Regards-Fi911 Support</p>  <!-- Closing Remarks -->
                    <p style="color:red;"><b>Note:</b> This is an automated message, please do not reply.</p>
                </body>
            </html>
            """
            send_email("Task Status Report", status_body)
            print("Status report email sent.")

            status_email_time = time.time()

        time.sleep(10)