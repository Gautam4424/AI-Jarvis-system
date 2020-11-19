from win32com.client import Dispatch
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
def speak(str):
    s=Dispatch("SAPI.SpVoice")
    s.speak(str)
def Take_command():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening....")
        # r.pause_threshold=1
        audio=r.listen(source)
        try:
            print("Recognizing....")
            query=r.recognize_google(audio,language="english")
            print(F"User said {query}")
            return query
        except Exception as e:
            print("Do not recognise the audio please speak again")
def wish_me():
    hour_time=datetime.datetime.now().hour
    if hour_time==6 or hour_time<12 :
        speak("Good Morning sir")
    elif hour_time>=12 and hour_time<19:
        speak("Good afternoon sir")
    elif hour_time>=19 or hour_time<22:
        speak("Good Evening Sir")
    else:
        speak("I hope your day is Good")

if __name__ == '__main__':
    wish_me()
    while True:
        query=Take_command().lower()
        if "wikipedia" in query:
            query.replace("wikipedia","")
            speak("According To Wikipedia")
            result=wikipedia.summary(query,sentences=1)
            speak(result)
            print(result)
        elif "who are you" in query:
            speak("My name is Jarvis Versin 2.0")
            speak("How may i help you sir")
        elif "open google" in query:
            webbrowser.open("www.google.com")
        elif "open youtube" in query:
            webbrowser.open("www.youtube.com")
        elif "open code with harry" in query:
            webbrowser.open("https://www.youtube.com/channel/UCeVMnSShP_Iviwkknt83cww")
        elif "open p" in query:
            path="C:\\Users\\Rohith Kumar\\Desktop\\Gautam\\PyCharm Community Edition 2019.3.1\\bin\\pycharm64.exe"
            os.startfile(path)








