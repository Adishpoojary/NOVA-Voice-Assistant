# NOVA – AI Voice Assistant with GUI

NOVA is a powerful desktop-based AI voice assistant that uses speech recognition, natural language processing, and a beautiful PyQt5 GUI to respond to user commands. It can open applications, fetch weather updates, read from Wikipedia, play media, send emails, and much more — all through voice.

---

## 🌟 Features

- 🎙️ **Voice-Controlled Interface** using Deepgram and Google TTS
- 💻 **System Monitoring Dashboard**: CPU, RAM, Disk, Battery usage
- 🌐 **Open Websites & Search Google/YouTube**
- 📚 **Search and Read Wikipedia**
- ☁️ **Get Real-Time Weather Data**
- 💌 **Send Emails via Gmail**
- ⏰ **Tell Time and Date**
- 🎵 **Media Control** (Play, Pause, Stop, Next, Previous)
- 🗃️ **Open Files, Folders, and Applications**
- 🖥️ **Real-Time Webcam Feed**
- ⚡ **System Commands** (Shutdown, Restart, Sleep)
- 💬 **Visual Feedback for Voice Commands**

---

## 🛠️ Tech Stack

- **Python 3**
- **PyQt5** – GUI and Dashboard
- **SpeechRecognition** + **Deepgram SDK** – For listening
- **gTTS + pygame** – For speaking
- **psutil** – System monitoring
- **OpenWeatherMap API** – Weather updates
- **Wikipedia API** – Knowledge lookup
- **smtplib** – Email sending
- **PyAutoGUI** – GUI automation
- **OpenCV** – Webcam stream
- **pycaw** – Volume control (Windows)
- **Threading/Multithreading** – For smooth performance

---

## ⚙️ Requirements

Install the dependencies using:

```bash
pip install -r requirements.txt
