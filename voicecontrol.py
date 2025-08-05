import sys
import threading
import time
import math
import socket
import os
import re
import json
import subprocess
import ctypes
import tempfile
import datetime

# GUI Imports
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QProgressBar, QFrame, QPushButton, QComboBox,
                             QGraphicsDropShadowEffect)
from PyQt5.QtCore import Qt, QTimer, QTime, QDate, QThread, pyqtSignal
from PyQt5.QtGui import (QImage, QPixmap, QPainter, QColor, QBrush, QPen,
                         QRadialGradient)

# Utility and Functionality Imports
import psutil
import cv2
import pyttsx3
import speech_recognition as sr
import webbrowser
import requests
import wikipedia
import feedparser
import pytz
from gtts import gTTS
import pygame  # Using pygame for reliable audio playback
import pyautogui
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


# Deepgram SDK
from deepgram import DeepgramClient, PrerecordedOptions, BufferSource

# --- CONFIGURATION & API KEYS ---
DG_API_KEY = '0060684156b491c391378b0199f10ee9ceab2268'
OPENWEATHER_API_KEY = "1f1f28772e8392767e9f49cdccd9fadb"
SENDER_EMAIL = "adishpoojary24@gmail.com"
SENDER_PASSWORD = "your_app_password" # IMPORTANT: Add your 16-character Gmail App Password here
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# --- MAPPINGS ---
WEBSITE_URLS = {
    "youtube": "https://www.youtube.com",
    "google": "https://www.google.com",
    "gmail": "https://mail.google.com",
    "github": "https://github.com",
}

# IMPORTANT: Using FULL PATHS for reliability. Check if these match your system.
APPLICATION_PATHS = {
    "notepad": "notepad.exe",
    "calculator": "calc.exe",
    "word": r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    "ms word": r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    "microsoft word": r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    "powerpoint": r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
    "excel": r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
    "chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    "cmd": "cmd.exe",
    "paint": "mspaint.exe",
    "file explorer": "explorer.exe",
    "command prompt": "cmd.exe",
}

FOLDER_COMMANDS = {
    "my pc": "shell:MyComputerFolder",
    "this pc": "shell:MyComputerFolder",
    "downloads": "shell:Downloads",
    "documents": "shell:Personal",
    "pictures": "shell:My Pictures",
    "music": "shell:My Music",
    "videos": "shell:My Video",
    "desktop": "shell:Desktop",
}

# --- GLOBAL VARIABLES ---
opened_processes = {}
pygame.mixer.init() # Initialize the pygame mixer

# --- CORE FUNCTIONS ---

def speak(text):
    """Converts text to speech using gTTS and plays it with pygame."""
    temp_path = ""
    try:
        print(f"NOVA: {text}")
        with tempfile.NamedTemporaryFile(delete=False, suffix='.mp3') as fp:
            temp_path = fp.name
        
        tts = gTTS(text=text, lang='en')
        tts.save(temp_path)
        
        pygame.mixer.music.load(temp_path)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            pygame.time.Clock().tick(10)
        
    except Exception as e:
        print(f"Error in speak function: {e}")
        engine = pyttsx3.init()
        engine.say(text)
        engine.runAndWait()
    finally:
        # Properly unload the music and remove the temp file
        pygame.mixer.music.unload()
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)

def listen():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1.5
        r.adjust_for_ambient_noise(source, duration=1)
        audio = r.listen(source)
    try:
        print("Recognizing with Deepgram...")
        wav_data = audio.get_wav_data()
        dg_client = DeepgramClient(api_key=DG_API_KEY)
        source = BufferSource(buffer=wav_data, mimetype="audio/wav")
        options = PrerecordedOptions(model="nova-2", smart_format=True)
        response = dg_client.listen.rest.v("1").transcribe_file(source, options)
        transcript = response.results.channels[0].alternatives[0].transcript
        print(f"User said: {transcript}")
        return transcript.lower()
    except Exception as e:
        print(f"Deepgram transcription error: {e}")
        return ""

# --- COMMAND FUNCTIONS ---

def get_time():
    now = datetime.datetime.now()
    return f"The current time is {now.strftime('%I:%M %p')}"

def get_date():
    now = datetime.datetime.now()
    return f"Today is {now.strftime('%A, %B %d, %Y')}"

def general_search(query):
    """Performs a general Google search."""
    try:
        search_url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
        webbrowser.open(search_url)
        return f"Here are the search results for {query}."
    except Exception as e:
        return f"Sorry, I couldn't perform the search. Error: {e}"
        
def search_youtube(query):
    """Opens YouTube and searches for a query."""
    try:
        search_url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
        webbrowser.open(search_url)
        return f"Searching YouTube for {query}."
    except Exception as e:
        return f"Sorry, I couldn't search YouTube. Error: {e}"

def search_wikipedia(query):
    try:
        speak("Searching Wikipedia...")
        results = wikipedia.summary(query, sentences=2)
        return "According to Wikipedia, " + results
    except Exception as e:
        return f"Sorry, I couldn't find information on {query}. Error: {e}"

def open_resource(target):
    target = target.strip(' .?!,')
    target_lower = target.lower()
    
    if target_lower in WEBSITE_URLS:
        webbrowser.open(WEBSITE_URLS[target_lower])
        return f"Opening {target.title()}."

    if target_lower in APPLICATION_PATHS:
        try:
            path = APPLICATION_PATHS[target_lower]
            proc = subprocess.Popen(path)
            opened_processes[target_lower] = proc
            return f"Opening {target.title()}."
        except FileNotFoundError:
            return f"Sorry, I couldn't find {target.title()} at '{path}'. Please check the file path."
        except Exception as e:
            return f"Sorry, I couldn't open {target.title()}: {e}"

    if target_lower in FOLDER_COMMANDS:
        try:
            os.startfile(FOLDER_COMMANDS[target_lower])
            return f"Opening {target.title()} folder."
        except Exception as e:
            return f"Sorry, I couldn't open {target.title()} folder: {e}"
            
    return f"Sorry, I don't know how to open '{target}'."

def close_application(app_name):
    app_key = app_name.lower().strip(' .?!,')
    
    if app_key in ["window", "active window", "youtube", "google"]:
        pyautogui.hotkey('alt', 'f4')
        return f"Closed the active window."
    
    proc = opened_processes.get(app_key)
    if proc and proc.poll() is None:
        try:
            proc.terminate()
            del opened_processes[app_key]
            return f"Closed {app_name}."
        except Exception as e:
            return f"Sorry, I couldn't close {app_name}: {e}"
    else:
        return f"{app_name} is not running or was not opened by me."

def get_weather(city):
    url = f"https://api.openweathermap.org/data/2.5/weather?q={city}&appid={OPENWEATHER_API_KEY}&units=metric"
    try:
        data = requests.get(url).json()
        if data.get("cod") != 200:
            return f"Sorry, I couldn't find weather for {city}."
        weather = data["weather"][0]["description"].capitalize()
        temp = data["main"]["temp"]
        return f"The weather in {city.title()} is {weather} with a temperature of {temp}°C."
    except Exception as e:
        return f"Sorry, an error occurred while fetching the weather: {e}"

def send_email(recipient, subject, message):
    if SENDER_PASSWORD == "mogl cksg hyio tjth":
        return "Email credentials are not set up."
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = recipient
        msg['Subject'] = subject
        msg.attach(MIMEText(message, 'plain'))
        
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        server.quit()
        return f"Email sent successfully to {recipient}."
    except Exception as e:
        return f"Sorry, I couldn't send the email: {e}"

def system_control(action):
    speak(f"Are you sure you want to {action} the computer? Please say yes or no.")
    confirmation = listen()
    if "yes" in confirmation:
        if action == "shutdown": os.system("shutdown /s /t 1")
        elif action == "restart": os.system("shutdown /r /t 1")
        elif action == "log off": os.system("shutdown /l")
        elif action == "sleep": ctypes.windll.PowrProf.SetSuspendState(0, 1, 0)
        return f"{action.capitalize()} initiated."
    else:
        return f"{action.capitalize()} cancelled."

def control_media(action):
    key_map = {"play": "playpause", "pause": "playpause", "stop": "stop", "next": "nexttrack", "previous": "prevtrack"}
    if action in key_map:
        pyautogui.press(key_map[action])
        return f"{action.capitalize()} command sent."
    return "Unknown media command."

# --- GUI CLASSES (No changes here) ---

class Dashboard(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("NOVA Dashboard")
        self.setStyleSheet("""
            QWidget { background-color: #101828; color: #fff; font-family: Arial; }
            QLabel { font-size: 16px; }
            QProgressBar { border: 1px solid #222; border-radius: 5px; background: #222; height: 18px; text-align: center;}
            QProgressBar::chunk { background-color: #00bfff; }
        """)
        self.init_ui()
        self.update_stats()
        self.update_time()
        self.update_network_info()
        self.cap = None
        self.start_webcam()

    def init_ui(self):
        main_layout = QHBoxLayout(self)

        left_panel = QVBoxLayout()
        self.cpu_label = QLabel("CPU Usage")
        self.cpu_bar = QProgressBar()
        self.ram_label = QLabel("RAM Usage")
        self.ram_bar = QProgressBar()
        self.disk_label = QLabel("Disk Usage")
        self.disk_bar = QProgressBar()
        self.battery_label = QLabel("Battery: N/A")
        
        left_panel.addWidget(self.cpu_label)
        left_panel.addWidget(self.cpu_bar)
        left_panel.addWidget(self.ram_label)
        left_panel.addWidget(self.ram_bar)
        left_panel.addWidget(self.disk_label)
        left_panel.addWidget(self.disk_bar)
        left_panel.addWidget(self.battery_label)
        left_panel.addStretch(1)

        left_frame = QFrame()
        left_frame.setLayout(left_panel)
        left_frame.setFixedWidth(220)

        center_panel = QVBoxLayout()
        self.orb_widget = OrbWidget()
        self.command_label = QLabel("Say your command")
        self.response_label = QLabel("Initializing...")
        
        self.command_label.setStyleSheet("font-style: italic; color: #ccc; font-size: 16px;")
        self.response_label.setStyleSheet("font-weight: bold; font-size: 18px;")
        
        center_panel.addStretch(1)
        center_panel.addWidget(self.orb_widget, alignment=Qt.AlignCenter)
        center_panel.addWidget(self.command_label, alignment=Qt.AlignCenter)
        center_panel.addWidget(self.response_label, alignment=Qt.AlignCenter)
        center_panel.addStretch(1)

        right_panel = QVBoxLayout()
        self.time_label = QLabel()
        self.network_label = QLabel()
        self.webcam_label = QLabel()
        self.webcam_label.setFixedSize(220, 165)
        self.webcam_label.setStyleSheet("background: #222; border-radius: 8px;")

        right_panel.addWidget(self.time_label, alignment=Qt.AlignTop | Qt.AlignHCenter)
        right_panel.addWidget(self.network_label, alignment=Qt.AlignTop | Qt.AlignHCenter)
        right_panel.addWidget(self.webcam_label, alignment=Qt.AlignCenter)
        right_panel.addStretch(1)

        right_frame = QFrame()
        right_frame.setLayout(right_panel)
        right_frame.setFixedWidth(240)

        main_layout.addWidget(left_frame)
        main_layout.addLayout(center_panel)
        main_layout.addWidget(right_frame)

        self.setLayout(main_layout)
        self.setMinimumSize(1100, 600)

    def update_stats(self):
        cpu = psutil.cpu_percent()
        ram = psutil.virtual_memory().percent
        disk = psutil.disk_usage('/').percent
        self.cpu_bar.setValue(int(cpu))
        self.ram_bar.setValue(int(ram))
        self.disk_bar.setValue(int(disk))
        
        try:
            battery = psutil.sensors_battery()
            if battery:
                plugged = '(Charging)' if battery.power_plugged else ''
                self.battery_label.setText(f"Battery: {battery.percent}% {plugged}")
            else:
                self.battery_label.setText("Battery: N/A")
        except (AttributeError, NotImplementedError):
            self.battery_label.setText("Battery: N/A")
            
        QTimer.singleShot(2000, self.update_stats)

    def update_time(self):
        now = QTime.currentTime().toString('HH:mm:ss')
        date = QDate.currentDate().toString('dddd, dd MMMM yyyy')
        self.time_label.setText(f"<b>{now}</b><br>{date}")
        QTimer.singleShot(1000, self.update_time)

    def update_network_info(self):
        try:
            ip = socket.gethostbyname(socket.gethostname())
            self.network_label.setText(f"IP Address: {ip}")
        except:
            self.network_label.setText("Network info unavailable")
        QTimer.singleShot(5000, self.update_network_info)

    def start_webcam(self, index=0):
        if self.cap: self.cap.release()
        self.cap = cv2.VideoCapture(index)
        if self.cap.isOpened():
            self.update_webcam()
        else:
            self.webcam_label.setText("Webcam not available")

    def update_webcam(self):
        if self.cap and self.cap.isOpened():
            ret, frame = self.cap.read()
            if ret:
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                h, w, ch = frame.shape
                qt_img = QImage(frame.data, w, h, ch * w, QImage.Format_RGB888)
                pixmap = QPixmap.fromImage(qt_img).scaled(self.webcam_label.width(), self.webcam_label.height(), Qt.KeepAspectRatio)
                self.webcam_label.setPixmap(pixmap)
        QTimer.singleShot(30, self.update_webcam)

    def closeEvent(self, event):
        if self.cap: self.cap.release()
        pygame.quit()
        QApplication.instance().quit()
        event.accept()

    def set_listening(self, is_listening):
        self.orb_widget.set_listening(is_listening)
    def show_response(self, text):
        self.response_label.setText(text)
    def show_command(self, text):
        if text:
            self.command_label.setText(f'Heard: "{text}"')
        else:
            self.command_label.setText("Listening...")

class OrbWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(200, 200)
        self.angle = 0
        self.listening = False
        self.pulse = 0
        self.pulse_dir = 1
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.animate)
        self.timer.start(30)

    def set_listening(self, listening):
        self.listening = listening
        self.update()

    def animate(self):
        self.angle = (self.angle + 1) % 360
        if self.listening:
            self.pulse += self.pulse_dir * 0.8
            if not 0 < self.pulse < 20: self.pulse_dir *= -1
        else:
            self.pulse = max(0, self.pulse - 1)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        w, h = self.width(), self.height()
        center = w // 2, h // 2
        
        orb_radius = 60 + int(self.pulse)
        grad = QRadialGradient(center[0], center[1], orb_radius)
        if self.listening:
            grad.setColorAt(0, QColor(0, 255, 127, 200))
            grad.setColorAt(0.5, QColor(0, 191, 255, 150))
            grad.setColorAt(1, QColor(0, 191, 255, 0))
        else:
            grad.setColorAt(0, QColor(58, 123, 213, 200))
            grad.setColorAt(1, QColor(58, 123, 213, 0))
            
        painter.setBrush(QBrush(grad))
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(center[0] - orb_radius, center[1] - orb_radius, orb_radius * 2, orb_radius * 2)

        painter.setBrush(QColor("#101828"))
        painter.setPen(QPen(QColor(0, 191, 255), 3))
        painter.drawEllipse(center[0] - 40, center[1] - 40, 80, 80)
        
        painter.setPen(QColor(255, 255, 255))
        font = self.font()
        font.setBold(True)
        font.setPointSize(12)
        painter.setFont(font)
        painter.drawText(self.rect(), Qt.AlignCenter, "NOVA")

# --- ASSISTANT THREAD ---

class AssistantThread(QThread):
    update_command = pyqtSignal(str)
    update_response = pyqtSignal(str)
    update_listening = pyqtSignal(bool)

    def run(self):
        speak("Initializing NOVA. How can I help you?")
        self.update_response.emit("Ready")
        
        while True:
            self.update_listening.emit(True)
            self.update_command.emit("")
            command = listen()
            self.update_listening.emit(False)
            
            if not command:
                continue

            self.update_command.emit(command)
            
            # --- NO WAKE WORD: Process command directly ---
            response = self.process_command(command)
            
            if response:
                self.update_response.emit(response)
                speak(response)
                
            time.sleep(1)
            
    def process_command(self, command):
        # --- File Commands ---
        if "save file" in command or "save project" in command:
            pyautogui.hotkey('ctrl', 's')
            return "Save command issued."

        # --- Search Commands ---
        elif command.startswith('search youtube for'):
            query = command.replace("search youtube for", "", 1).strip()
            return search_youtube(query)
            
        elif command.startswith("search for"):
            query = command.replace("search for", "", 1).strip()
            return general_search(query)
            
        elif 'wikipedia' in command:
            query = command.replace('wikipedia', '').replace('search for', '').strip(' .?!,')
            return search_wikipedia(query)

        # --- General Commands ---
        elif command.startswith('open '):
            target = command.replace('open', '', 1).strip()
            return open_resource(target)

        elif command.startswith('close '):
            target = command.replace('close', '', 1).strip()
            return close_application(target)
        
        elif 'weather in' in command:
            city = command.split('in', 1)[-1].strip(' .?!,')
            return get_weather(city)

        elif 'time' in command:
            return get_time()

        elif 'date' in command:
            return get_date()

        elif command.startswith('type'):
            text_to_type = command.replace('type', '', 1).strip()
            pyautogui.typewrite(text_to_type)
            return f"Typing: {text_to_type}"
        
        elif any(x in command for x in ["play", "pause", "stop", "next", "previous"]):
            action = command.split(" ")[0]
            return control_media(action)
        
        elif "send email" in command:
            try:
                speak("To whom should I send the email?")
                recipient = listen()
                speak("What is the subject?")
                subject = listen()
                speak("What should be the message?")
                message = listen()
                return send_email(recipient, subject, message)
            except Exception as e:
                return f"Could not complete the email request. {e}"

        elif any(x in command for x in ["shutdown", "restart", "log off", "sleep"]):
            action = command.split()[-1]
            return system_control(action)

        elif "what is your name" in command or "who are you" in command:
            return "My name is NOVA, your personal AI assistant."
            
        else:
            return "I don't understand that command."

# --- MAIN EXECUTION ---
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    dashboard = Dashboard()
    dashboard.show()
    
    assistant_thread = AssistantThread()
    assistant_thread.update_command.connect(dashboard.show_command)
    assistant_thread.update_response.connect(dashboard.show_response)
    assistant_thread.update_listening.connect(dashboard.set_listening)
    
    app.aboutToQuit.connect(assistant_thread.terminate)
    
    assistant_thread.start()
    
    sys.exit(app.exec_())
