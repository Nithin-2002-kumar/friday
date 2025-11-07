import speech_recognition as sr
import pyttsx3
import pywhatkit
import datetime
import os
import subprocess
import psutil
import platform
import pyautogui
import shutil
import winreg
import json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ---------- SETUP ----------
engine = pyttsx3.init()
recognizer = sr.Recognizer()
CACHE_FILE = "apps_cache.json"
TASK_FILE = "tasks.json"

# ----- EMAIL CONFIG -----
YOUR_EMAIL = "youremail@gmail.com"  # <-- your Gmail address
YOUR_APP_PASSWORD = "your_app_password_here"  # <-- Gmail App password

def speak(text):
    print("friday:", text)
    engine.say(text)
    engine.runAndWait()

def listen():
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)
    try:
        command = recognizer.recognize_google(audio)
        print(f"You said: {command}")
        return command.lower()
    except:
        speak("Sorry, I didn’t catch that.")
        return ""

# ---------------- TASK MANAGEMENT ----------------
def load_tasks():
    if os.path.exists(TASK_FILE):
        with open(TASK_FILE, "r") as f:
            return json.load(f)
    return []

def save_tasks(tasks):
    with open(TASK_FILE, "w") as f:
        json.dump(tasks, f, indent=2)

def add_task(task):
    tasks = load_tasks()
    tasks.append(task)
    save_tasks(tasks)
    speak(f"Task '{task}' added.")
    return f"Task '{task}' added."

def list_tasks():
    tasks = load_tasks()
    if not tasks:
        speak("You have no tasks.")
        return "No tasks found."
    speak(f"You have {len(tasks)} tasks.")
    result = "\n".join([f"{i+1}. {t}" for i, t in enumerate(tasks)])
    return result

def delete_task(task):
    tasks = load_tasks()
    for t in tasks:
        if task in t:
            tasks.remove(t)
            save_tasks(tasks)
            speak(f"Deleted task '{task}'.")
            return f"Deleted task '{task}'."
    speak("Task not found.")
    return "Task not found."

def clear_tasks():
    save_tasks([])
    speak("All tasks cleared.")
    return "All tasks deleted."

# ---------------- EMAIL ----------------
def send_email(receiver, subject, message):
    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(YOUR_EMAIL, YOUR_APP_PASSWORD)

        email = MIMEMultipart()
        email["From"] = YOUR_EMAIL
        email["To"] = receiver
        email["Subject"] = subject
        email.attach(MIMEText(message, "plain"))

        server.sendmail(YOUR_EMAIL, receiver, email.as_string())
        server.quit()

        speak(f"Email sent to {receiver}.")
        return f"Email successfully sent to {receiver}."
    except Exception as e:
        speak("Failed to send email.")
        return f"Error: {e}"

def handle_send_email():
    speak("Who should I send the email to?")
    receiver = listen()
    if "at the rate" in receiver:
        receiver = receiver.replace(" at the rate ", "@")
    receiver = receiver.replace(" ", "")  # remove spaces

    speak("What is the subject?")
    subject = listen()

    speak("What is the message?")
    message = listen()

    return send_email(receiver, subject, message)

# ---------- APP DISCOVERY ----------
def get_installed_apps():
    """Scans Windows registry for installed apps"""
    apps = {}
    reg_paths = [
        r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall",
        r"SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall",
    ]

    for reg_path in reg_paths:
        try:
            reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
            for i in range(0, winreg.QueryInfoKey(reg_key)[0]):
                try:
                    subkey_name = winreg.EnumKey(reg_key, i)
                    subkey = winreg.OpenKey(reg_key, subkey_name)
                    name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                    try:
                        install_path = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                    except FileNotFoundError:
                        install_path = ""
                    apps[name.lower()] = install_path
                except (FileNotFoundError, OSError, IndexError):
                    continue
        except Exception:
            continue
    return apps

def cache_apps(apps):
    with open(CACHE_FILE, "w") as f:
        json.dump(apps, f)

def load_cached_apps():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r") as f:
            return json.load(f)
    return {}

def find_app_path(app_name):
    """Find app path dynamically using cache, registry, and Start Menu"""
    app_name = app_name.lower().strip()
    cached_apps = load_cached_apps()

    # 1. Check cache
    for name, path in cached_apps.items():
        if app_name in name:
            return path

    # 2. Search registry
    apps = get_installed_apps()
    for name, path in apps.items():
        if app_name in name:
            if path:
                cached_apps[name] = path
                cache_apps(cached_apps)
                return path

    # 3. Search Start Menu shortcuts
    start_menu_paths = [
        os.path.expandvars(r"%ProgramData%\Microsoft\Windows\Start Menu\Programs"),
        os.path.expandvars(r"%AppData%\Microsoft\Windows\Start Menu\Programs"),
    ]

    for base in start_menu_paths:
        for root, dirs, files in os.walk(base):
            for file in files:
                if app_name in file.lower() and file.endswith(".lnk"):
                    shortcut = os.path.join(root, file)
                    cached_apps[app_name] = shortcut
                    cache_apps(cached_apps)
                    return shortcut

    return None

# ---------- APP CONTROL ----------
def open_application(app_name):
    path = find_app_path(app_name)
    if path:
        speak(f"Opening {app_name}")
        try:
            os.startfile(path)
        except:
            subprocess.Popen(f"start {path}", shell=True)
        return f"Opened {app_name}."
    else:
        speak(f"I couldn’t find {app_name} installed.")
        return f"{app_name} not found."

def close_application(app_name):
    app_name = app_name.lower().strip()
    closed = False
    for proc in psutil.process_iter(['pid', 'name']):
        if app_name in proc.info['name'].lower():
            psutil.Process(proc.info['pid']).terminate()
            closed = True
    if closed:
        speak(f"Closed {app_name}.")
        return f"{app_name} closed."
    else:
        speak(f"No running process found for {app_name}.")
        return f"Couldn't find {app_name} running."

# ---------- SYSTEM INFO ----------
def get_system_info():
    info = (
        f"System: {platform.system()} {platform.release()}\n"
        f"Processor: {platform.processor()}\n"
        f"CPU Usage: {psutil.cpu_percent()}%\n"
        f"RAM Usage: {psutil.virtual_memory().percent}%\n"
        f"Battery: {psutil.sensors_battery().percent if psutil.sensors_battery() else 'N/A'}%"
    )
    return info

# ---------- COMMAND HANDLER ----------
def run_command(command):
    response = ""

    # --- Email ---
    if "send email" in command:
        response = handle_send_email()

    # --- Tasks ---
    elif "add task" in command:
        task = command.replace("add task", "").strip()
        response = add_task(task)

    elif "show tasks" in command or "list tasks" in command:
        response = list_tasks()

    elif "delete task" in command:
        task = command.replace("delete task", "").strip()
        response = delete_task(task)

    elif "clear tasks" in command:
        response = clear_tasks()

    elif "open" in command:
        app = command.replace("open", "").strip()
        response = open_application(app)

    elif "close" in command:
        app = command.replace("close", "").strip()
        response = close_application(app)

    elif "time" in command:
        time = datetime.datetime.now().strftime("%I:%M %p")
        response = f"The time is {time}"
        speak(response)

    elif "date" in command:
        date = datetime.datetime.now().strftime("%A, %B %d, %Y")
        response = f"Today is {date}"
        speak(response)

    elif "search" in command:
        query = command.replace("search", "")
        speak(f"Searching Google for {query}")
        pywhatkit.search(query)

    elif "system info" in command:
        info = get_system_info()
        speak("Here are your system details.")
        response = info

    elif "code" in command:
        speak("Opening Visual Studio Code")
        subprocess.Popen("code")  # Assumes VS Code in PATH
        # File operations
    elif "create folder" in command:
        name = command.replace("create folder", "").strip()
        if name:
            os.mkdir(name)
            response = f"Folder '{name}' created."
            speak(response)
        else:
            speak("Please say the folder name.")

    elif "delete folder" in command:
        name = command.replace("delete folder", "").strip()
        if os.path.exists(name):
            shutil.rmtree(name)
            response = f"Folder '{name}' deleted."
            speak(response)
        else:
            speak("Folder not found.")

    elif "create file" in command:
        name = command.replace("create file", "").strip()
        if name:
            with open(name, "w") as f:
                f.write("")
            response = f"File '{name}' created."
            speak(response)
        else:
            speak("Please say the file name.")

    elif "delete file" in command:
        name = command.replace("delete file", "").strip()
        if os.path.exists(name):
            os.remove(name)
            response = f"File '{name}' deleted."
            speak(response)
        else:
            speak("File not found.")

    if "open chrome" in command:
        speak("Opening Chrome")
        os.startfile("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe")

    elif "open notepad" in command:
        speak("Opening Notepad.")
        subprocess.Popen("notepad.exe")

    elif "open folder" in command:
        path = command.replace("open folder", "").strip()
        if os.path.exists(path):
            os.startfile(path)
            speak(f"Opening folder {path}")
        else:
            speak("Folder not found.")

    elif "screenshot" in command:
        filename = f"screenshot_{datetime.datetime.now().strftime('%H-%M-%S')}.png"
        pyautogui.screenshot(filename)
        response = f"Screenshot saved as {filename}"
        speak(response)

    elif "volume up" in command:
        for _ in range(5): pyautogui.press("volumeup")
        speak("Volume increased.")

    elif "volume down" in command:
        for _ in range(5): pyautogui.press("volumedown")
        speak("Volume decreased.")

    elif "mute" in command:
        pyautogui.press("volumemute")
        speak("Volume muted.")

    elif "shutdown" in command:
        speak("Shutting down the system.")
        os.system("shutdown /s /t 1")

    elif "restart" in command:
        speak("Restarting system.")
        os.system("shutdown /r /t 1")

    elif "log off" in command:
        speak("Logging off.")
        os.system("shutdown /l")

    elif "exit" in command or "quit" in command:
        speak("Goodbye!")
        exit()

    else:
        speak("Sorry, I don't know that command.")
        response = "Command not recognized."

    return response
