from vosk import Model, KaldiRecognizer
from dateutil.parser import parse
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import os
import pyaudio
import json
import pyttsx3
import webbrowser
import keyboard
import speech_recognition as sr
import pyjokes
import ctypes
"""
Version: 2021-10-17

This program uses the Vosk library.
Voice recognition with Python.
Detects voice commands and performs specific operations.
"""

#model for voice recognition
model = Model("model")

#custom invalid format error class
class InvalidFormatError(Exception):
    pass

class HandleAudio:
    def __init__(self, text):
        self.text = text

    def parse_tasks(self):
        if self.text != " " and self.text != "":
            #commands for the voice assistant
            if "open" in self.text:
                if "memorial mail" in self.text or "memorial mayo" in self.text:
                    webbrowser.open("https://mail.google.com/mail/u/0/#inbox")
                elif "memorial on line" in self.text or "memorial online" in self.text:
                    webbrowser.open("https://online.mun.ca/d2l/home")
                elif "my website" in self.text:
                    webbrowser.open("https://www.bwthorne.com/")
                elif "reminder application" in self.text or "reminder app" in self.text:
                    os.startfile("C:/Users/brend/PycharmProjects/Alarm/ReminderApp.exe")
                return True

            if ("change" in self.text or "set" in self.text or "turn" in self.text) and "volume" in self.text:
                if "zero" in self.text:
                    set_audio("0")
                elif "twenty five" in self.text:
                    set_audio("25")
                elif "fifty" in self.text:
                    set_audio("50")
                elif "seventy five" in self.text:
                    set_audio("75")
                elif "one hundred" in self.text:
                    set_audio("100")
                return True

            if ("read" in self.text or "what are" in self.text or "list" in self.text) and "reminders" in self.text:
                set_audio("25") #setting audio to 25 since tts is loud
                with open("C:/RemindersData/Reminders.txt", "r+") as f:
                    lines = f.readlines()
                    for line in lines:
                        text_message = line[line.find("(") + 1:line.find(")")]
                        py_speak(text_message)
                    f.close()
                return True

            if "stop listening" in self.text:
                py_speak("Done listening")
                return False

            if "tell" in self.text and "joke" in self.text:
                py_speak(pyjokes.get_joke())
                return True

            if "lock" in self.text and "computer" in self.text:
                ctypes.windll.user32.LockWorkStation()
                return True

            if "sleep" in self.text and "computer" in self.text:
                py_speak("Putting computer to sleep")
                os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
                return True
        return True


class HandleReminders:
    def __init__(self, reminder_time, reminder_message):
        self.reminder_time = reminder_time
        self.reminder_message = reminder_message

    def format_reminder_time(self):
        #time will always be this string slice
        actual_reminder_time = self.reminder_time[11:]
        #date will always be this string slice
        actual_date = self.reminder_time[:10]
        final_time = ""

        if "a.m." in self.reminder_time:
            reminder_time_formatted = actual_reminder_time.replace("a.m.", "").rstrip() + ":00" + ":00" + " AM"
            date_time_formatted = actual_date.replace(" ", "")
            final_time = date_time_formatted + " " + reminder_time_formatted

        elif "p.m." in self.reminder_time:
            reminder_time_formatted = actual_reminder_time.replace("p.m.", "").rstrip() + ":00" + ":00" + " PM"
            date_time_formatted = actual_date.replace(" ", "")
            final_time = date_time_formatted + " " + reminder_time_formatted
        return final_time

    def write_to_file(self, method):
        f = open("C:/RemindersData/Reminders.txt", method)
        #enclosing message in parentheses and time in brackets
        f.write("(" + self.reminder_message + ")" + " " + "[" + self.format_reminder_time() + "]")
        f.write("\n")
        f.close()
        py_speak("Reminder added")

    def handle_write(self):
        if os.stat("C:/RemindersData/Reminders.txt").st_size == 0:
            self.write_to_file("w")
        else:
            self.write_to_file("a")


def set_audio(value):
    #get default audio device using PyCAW
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    #get current volume
    currentVolumeDb = volume.GetMasterVolumeLevel()

    """
    These values seem to be specific to my computer, as other users of PyCAW have reported
    different decibel values corresponding to different master volume levels
    """
    if value == "0":
        volume.SetMasterVolumeLevel(-50.0, None) #sets volume to 0 (-50.0 corresponds to decibel value)
    elif value == "25":
        volume.SetMasterVolumeLevel(-16.0, None) #sets volume to 25
    elif value == "50":
        volume.SetMasterVolumeLevel(-6.25, None) #sets volume to 50
    elif value == "75":
        volume.SetMasterVolumeLevel(-0.5, None) #sets volume to 75
    elif value == "100":
        volume.SetMasterVolumeLevel(5.0, None) #sets volume to 100


def collect_results(recognizer):
    result = recognizer.Result()
    wjdata = json.loads(result)
    return wjdata["text"] #return text format of json data


def py_speak(message):
    engine = pyttsx3.init(driverName="sapi5")
    voices = engine.getProperty("voices")
    engine.setProperty('voice', voices[0].id)
    rate = engine.getProperty("rate")
    engine.setProperty("rate", 130) #speed at which the bot will speak
    engine.say(message)
    engine.runAndWait()
    engine.stop()


def get_audio_from_google():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source)
        print("Listening")
        audio = r.listen(source)
        said = ""
    try:
        said = r.recognize_google(audio)
        print(said)

    except sr.UnknownValueError as e:
        py_speak("Could not understand.")
        print("Could not understand audio; " + str(e))

    except sr.RequestError as e:
        py_speak("Could not request results.")
        print("Could not request results from Google; " + str(e))

    return said


def has_numbers(input):
    #return true if input has numbers else return false
    return any(char.isdigit() for char in input)


def main():
    listening = False
    listen_to_reminder = False

    rec = KaldiRecognizer(model, 16000)
    p = pyaudio.PyAudio()

    stream = p.open(format=pyaudio.paInt16, channels=1, rate=16000, input=True, frames_per_buffer=8000)

    while True:
        data = stream.read(4000, exception_on_overflow=False)
        text_data = ""

        if rec.AcceptWaveform(data):
            text_data = collect_results(rec)
            print(text_data)

        if "wake up python" in text_data or "wake up" in text_data:
            py_speak("Listening")

            listen_rec = KaldiRecognizer(model, 16000)
            listening = True

        while listening:
            text = ""
            data_listen = stream.read(4000, exception_on_overflow=False)

            if listen_rec.AcceptWaveform(data_listen):
                text = collect_results(listen_rec)
                print(text)

            audio = HandleAudio(text)
            if not audio.parse_tasks():
                listening = False
            if "set a reminder" in text:
                listen_to_reminder = True
                listening = False

        count = 0
        while listen_to_reminder:
            while count < 1:
                py_speak("What would you like to be reminded of?")
                reminder_msg = get_audio_from_google()
                py_speak("You said " + reminder_msg + ". Is this correct?")
                answer_one = get_audio_from_google()
                if answer_one == "yes":
                    count += 1

            while count < 2:
                year_month_day = ""
                time = ""
                while True:
                    try:
                        py_speak("What date would you like to be reminded?")
                        reminder_date = get_audio_from_google()
                        year_month_day = parse(reminder_date).strftime("%Y-%m-%d") # parsing date into given format Y-m-d
                        break
                    except Exception as e:
                        py_speak(e + ".")
                        print(e)

                while True:
                    try:
                        py_speak("What time would you like to be reminded?")
                        time = get_audio_from_google()
                        if has_numbers(time) and "a.m" in time or "p.m." in time:
                            break
                        else:
                            raise InvalidFormatError("Invalid format")
                    except InvalidFormatError as exp:
                        py_speak("Invalid format.")
                        print(exp)

                reminder_time = year_month_day + " " + time
                py_speak("You said " + reminder_time + ". Is this correct?")
                answer_two = get_audio_from_google()

                if answer_two == "yes":
                    count += 1
            reminders = HandleReminders(reminder_time, reminder_msg)
            reminders.handle_write()
            listen_to_reminder = False


if __name__ == "__main__":
    main()
