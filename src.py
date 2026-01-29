import os
import json
import datetime
import webbrowser

import pyaudio
import pyttsx3
from vosk import Model, KaldiRecognizer

# ---------- COM & VOLUME ----------
import comtypes
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# ---------- BRIGHTNESS ----------
import screen_brightness_control as sbc


# ================= COM INITIALIZATION =================
comtypes.CoInitialize()


# ================= TEXT TO SPEECH =================
engine = pyttsx3.init()
engine.setProperty("rate", 165)

def speak(text):
    engine.say(text)
    engine.runAndWait()


# ================= VOLUME SETUP (SAFE) =================
try:
    speakers = AudioUtilities.GetSpeakers()
    interface = speakers.Activate(
        IAudioEndpointVolume._iid_,
        CLSCTX_ALL,
        None
    )
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    def volume_up():
        v = volume.GetMasterVolumeLevelScalar()
        volume.SetMasterVolumeLevelScalar(min(v + 0.1, 1.0), None)

    def volume_down():
        v = volume.GetMasterVolumeLevelScalar()
        volume.SetMasterVolumeLevelScalar(max(v - 0.1, 0.0), None)

    def mute():
        volume.SetMute(1, None)

    def unmute():
        volume.SetMute(0, None)

    volume_available = True
except Exception as e:
    print("Volume control unavailable:", e)
    volume_available = False


# ================= SPEECH RECOGNITION =================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VOSK_PATH = os.path.join(BASE_DIR, "vosk-model-small-en-us-0.15")

model = Model(VOSK_PATH)
recognizer = KaldiRecognizer(model, 16000)

mic = pyaudio.PyAudio()
stream = mic.open(
    format=pyaudio.paInt16,
    channels=1,
    rate=16000,
    input=True,
    frames_per_buffer=4096
)

stream.start_stream()
speak("Offline voice assistant started")


# ================= MAIN LOOP =================
while True:
    data = stream.read(4096, exception_on_overflow=False)

    if recognizer.AcceptWaveform(data):
        result = json.loads(recognizer.Result())
        command = result.get("text", "").lower()

        if not command:
            continue

        print("Command:", command)

        # ----- APPLICATIONS -----
        if "open chrome" in command:
            speak("Opening Chrome")
            os.system("start chrome")

        elif "open notepad" in command:
            speak("Opening Notepad")
            os.system("notepad")

        elif "open calculator" in command:
            speak("Opening Calculator")
            os.system("calc")

        elif "open explorer" in command:
            speak("Opening File Explorer")
            os.system("explorer")

        # ----- VOLUME -----
        elif ("increase volume" in command or "volume up" in command) and volume_available:
            speak("Increasing volume")
            volume_up()

        elif ("decrease volume" in command or "volume down" in command) and volume_available:
            speak("Decreasing volume")
            volume_down()

        elif "mute" in command and volume_available:
            speak("Muted")
            mute()

        elif "unmute" in command and volume_available:
            speak("Unmuted")
            unmute()

        # ----- BRIGHTNESS -----
        elif "increase brightness" in command:
            try:
                sbc.set_brightness(min(sbc.get_brightness()[0] + 10, 100))
                speak("Increasing brightness")
            except:
                speak("Brightness control not supported")

        elif "decrease brightness" in command:
            try:
                sbc.set_brightness(max(sbc.get_brightness()[0] - 10, 0))
                speak("Decreasing brightness")
            except:
                speak("Brightness control not supported")

        # ----- TIME & DATE -----
        elif "time" in command:
            speak(datetime.datetime.now().strftime("Time is %H %M"))

        elif "date" in command:
            speak(datetime.datetime.now().strftime("Today is %d %B %Y"))

        # ----- WEB -----
        elif "open youtube" in command:
            speak("Opening YouTube")
            webbrowser.open("https://youtube.com")

        elif "open google" in command:
            speak("Opening Google")
            webbrowser.open("https://google.com")

        # ----- SYSTEM -----
        elif "shutdown" in command:
            speak("Shutting down")
            os.system("shutdown /s /t 5")

        elif "restart" in command:
            speak("Restarting")
            os.system("shutdown /r /t 5")

        elif "lock system" in command:
            speak("Locking system")
            os.system("rundll32.exe user32.dll,LockWorkStation")

        elif "exit" in command or "stop assistant" in command:
            speak("Goodbye")
            break


# ================= CLEANUP =================
stream.stop_stream()
stream.close()
mic.terminate()
comtypes.CoUninitialize()
