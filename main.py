from win32com.client import Dispatch
import datetime

def speak(audio):
    speaking = Dispatch('SAPI.Spvoice')
    speaking.speak(audio)
# speak("Hello Swastik Sir")

def time():
    t_now = datetime.datetime.now().strftime('%H:%M:%S')
    # print(t_now)
    speak("The current time")
    speak(t_now)
# time()

def greet():
    t_hour = datetime.datetime.now().hour
    print(t_hour)