from win32com.client import Dispatch
import datetime
import pytz
import speech_recognition as sr

def speak(audio):
    speaking = Dispatch('SAPI.Spvoice')
    speaking.speak(audio)
speak("Hello Swastik Sir")

def time():
    t_now = datetime.datetime.now().strftime('%H:%M:%S')
    # print(t_now)
    speak("The current time")
    speak(t_now)
# time()

def greet():
    t_hour = datetime.datetime.now().hour
    # print(t_hour)
    if 24 > t_hour <= 4:
        speak("Have a pleasant night Sir.")
    elif 12 > t_hour > 4:
        speak("Good Morning Sir")
    elif 18 > t_hour >= 12:
        speak("Good Afternoon Sir")
    else:
        speak("Good Evening Sir")
        speak("Hope you enjoyed your day!")

# greet()

def date():
    t_date = datetime.datetime.now(tz=pytz.timezone('Asia/Kolkata'))
    print(t_date.strftime('%d %B, %Y'))
    speak(t_date.strftime('%d %B, %Y'))
# date()

def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("JARVIS at your service.")
        print("Command me sir")
        r.pause_threshold = 1
        command = r.listen(source)
    try:
        print("Recognizing your request.....")
        recognized = r.recognize_google(command, 'en-in')
        print(recognized)
    except Exception as e:
        print(e)
        statement = "Pardon Sir. Couldn't process your request."
        print(statement)
        speak(statement)
        return None
    return recognized
# take_command()