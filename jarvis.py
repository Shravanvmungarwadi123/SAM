import subprocess
import wolframalpha
import pyttsx3
# import tkinter
import json
import random
# import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import sys
import winshell
import pyjokes
# import feedparser
import smtplib
import ctypes
import time
import requests
from requests import get
import pyautogui
# import shutil
from twilio.rest import Client
from clint.textui import progress
# import ecapture as ec
# from bs4 import BeautifulSoup
# import win32com.client as wincl
from urllib.request import urlopen
import instaloader
#import pywhatkit as kit
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import QTimer, QTime, QDate, Qt
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from jarvisUi import Ui_MainWindow



engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voices', voices[0].id)


def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    tt = time.strftime("%I:%M %p")

    if 0 <= hour < 12:
        speak(f"Good Morning Sir !, Its {tt}")

    elif 12 <= hour < 18:
        speak(f"Good Afternoon Sir !, Its {tt}")

    else:
        speak(f"Good Evening Sir !, Its {tt}")
    speak("I am your Assistant SAM")

class MainThread(QThread):
    def __init__(self):
        super(MainThread, self).__init__()

    def run(self):
        self.TaskExecution()



    def takeCommand(self):
        r = sr.Recognizer()

        with sr.Microphone() as source:

            print("Listening...")
            r.pause_threshold = 1
            audio = r.listen(source, timeout=5, phrase_time_limit=8)

        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language='en-in')
            print(f"User said: {query}\n")

        except Exception as e:
            print(e)
            print("Unable to Recognize your voice.")
            return "None"

        return query


    def TaskExecution(self):
        wishMe()
        assname="SAM"
        while True:
            self.query = self.takeCommand().lower()
            #found by SGB-1
            '''
            if 'wikipedia' in self.query:
                try:
                    speak('Searching Wikipedia...')
                    self.query = self.query.replace("Wikipedia", "")
                    results = wikipedia.summary(self.query, sentences=3)
                    speak("According to Wikipedia")
                    speak(results)
                except Exception as e:
                    print(e)
                    speak("I am not able to find please search for other information.")
            '''
            if 'open youtube' in self.query:
                speak("Here you go to Youtube\n")
                webbrowser.open("youtube.com")
                
            #elif 'who is ' in self.query:
             #   speak("")

            elif 'open whatsapp' in self.query:
                speak("Here you go to Whatsapp\n")
                webbrowser.open("web.whatsapp.com")

            elif 'open google' in self.query:
                speak("Here you go to Google\n")
                webbrowser.open("google.com")

            elif 'stack overflow' in self.query:
                speak("Here you go to Stack Over flow.Happy coding")
                webbrowser.open("stackoverflow.com")

            elif 'play music' in self.query or "play song" in self.query:
                speak("Here you go with music")
                music_dir = "D:\\Music"
                songs = os.listdir(music_dir)
                rd = random.choice(songs)
                os.startfile(os.path.join(music_dir, rd))

            elif 'the time' in self.query:
                strTime = datetime.datetime.now().strftime("%I:%M %p")
                speak(f"Sir, the time is {strTime}")

            elif 'open opera' in self.query:
                codePath = r"C:\\Users\\GAURAV\\AppData\\Local\\Programs\\Opera\\launcher.exe"
                os.startfile(codePath)

            elif 'email to gaurav' in self.query:
                try:
                    speak("What should I say?")
                    content = self.takeCommand()
                    to = "Receiver email address"
                    sendEmail(to, content)
                    speak("Email has been sent !")
                except Exception as e:
                    print(e)
                    speak("I am not able to send this email")

            elif 'send a mail' in self.query:
                try:
                    speak("What should I say?")
                    content = self.takeCommand()
                    speak("whom should i send")
                    to = input()
                    sendEmail(to, content)
                    speak("Email has been sent !")
                except Exception as e:
                    print(e)
                    speak("I am not able to send this email")

            elif 'how are you' in self.query:
                speak("I am fine, Thank you")
                speak("How are you, Sir")

            elif 'fine' in self.query or "good" in self.query:
                speak("It's good to know that your fine")

            elif "change my name to" in self.query:
                query = self.query.replace("change my name to", "")
                assname = query

            elif "change name" in self.query:
                speak("What would you like to call me, Sir ")
                assname = self.takeCommand()
                speak("Thanks for naming me")
                
            elif "what's your name" in self.query or "What is your name" in self.query:
                try:
                    speak(f"My friends call me {assname}")
                except Exception as e:
                    print(e)
                    speak("I am not able to understand")
               

            elif 'exit' in self.query or 'you can go now' in self.query:
                speak("Thanks for giving me your time")
                exit()

            elif "who made you" in self.query or "who created you" in self.query:
                speak("I have been created by Subramanya.")

            elif 'joke' in self.query:
                speak(pyjokes.get_joke())

            elif "calculate" in self.query:

                app_id = "EPHWP8-W46635H8VG"
                client = wolframalpha.Client(app_id)
                indx = self.query.lower().split().index('calculate')
                query = self.query.split()[indx + 1:]
                res = client.query(' '.join(query))
                answer = next(res.results).text
                print("The answer is " + answer)
                speak("The answer is " + answer)

            elif 'search' in self.query or 'play' in self.query:

                query = self.query.replace("search", "")
                query = self.query.replace("play", "")
                webbrowser.open(query)

            elif "who i am" in self.query:
                speak("If you talk then definitely your human.")

            elif "why you came to world" in self.query:
                speak("Thanks to Subramanya. further It's a secret")

            elif 'power point presentation' in self.query:
                speak("opening Power Point presentation")
                power = r"C:\\Users\\subra\\Desktop\\Minor Project\\Presentation\\Voice Assistant.pptx"
                os.startfile(power)

            elif 'is love' in self.query:
                speak("It is 7th sense that destroy all other senses")

            elif "who are you" in self.query:
                speak("I am your virtual assistant created by Subramanya")

            elif 'reason for you' in self.query:
                speak("I was created as a Minor project by Mister Subramanya ")

            elif 'change background' in self.query:
                ctypes.windll.user32.SystemParametersInfoW(20, 0, "Location of wallpaper", 0)
                speak("Background changed successfully")

            elif 'open bluestack' in self.query:
                appli = r"C:\\ProgramData\\BlueStacks\\Client\\Bluestacks.exe"
                os.startfile(appli)

            elif 'news' in self.query:

                try:
                    jsonObj = urlopen('''http://newsapi.org/v2/top-headlines?country=in&category=technology&apiKey=c3285e34e7cd41de9452043e67f873f6''')
                    data = json.load(jsonObj)
                    i = 1

                    speak('here are some top news Sir.')

                    for item in data['articles']:
                        speak(str(i) + '. ' + item['title'] + '\n')
                        speak(item['description'] + '\n')
                        #speak(str(i) + '. ' + item['title'] + '\n')
                        i += 1
                except Exception as e:
                    print(str(e))

            elif 'lock system' in self.query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()

            elif 'shutdown system' in self.query:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')

            elif 'empty recycle bin' in self.query:
                winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
                speak("Recycle Bin Recycled")

            elif "don't listen" in self.query or "stop listening" in self.query:
                speak("for how much time you want to stop jarvis from listening commands")
                a = int(self.takeCommand())
                time.sleep(a)
                print(a)

            elif "where is" in self.query:
                query = self.query.replace("where is", "")
                location = query
                speak("Hold on sir, I will show you where " + location + " is.")
                speak(location)
                webbrowser.open("https://www.google.nl / maps / place/" + location + "")

            # elif "camera" in query or "take a photo" in query:
            # ec.capture(0, "Jarvis Camera ", "img.jpg")

            elif "restart" in self.query:
                subprocess.call(["shutdown", "/r"])

            elif "hibernate" in self.query or "sleep" in self.query:
                speak("Hibernating")
                subprocess.call("shutdown / h")

            elif "log off" in self.query or "sign out" in self.query:
                speak("Make sure all the application are closed before sign-out")
                time.sleep(5)
                subprocess.call(["shutdown", "/l"])

            elif "write a note" in self.query:
                speak("What should i write, sir")
                note = self.takeCommand()
                file = open('jarvis.txt', 'w')
                speak("Sir, Should i include date and time")
                snfm = self.takeCommand()
                if 'yes' in snfm or 'sure' in snfm:
                    strTime = datetime.datetime.now().strftime("% H:% M:% S")
                    file.write(strTime)
                    file.write(" :- ")
                    file.write(note)
                else:
                    file.write(note)

            elif "show note" in self.query:
                speak("Showing Notes")
                file = open("jarvis.txt", "r")
                print(file.read())
                speak(file.read(6))
          
                
            elif "update assistant" in self.query:
                speak("After downloading file please replace this file with the downloaded one")
                url = '# url after uploading file'
                r = requests.get(url, stream=True)

                with open("Voice.py", "wb") as Pypdf:

                    total_length = int(r.headers.get('content-length'))

                    for ch in progress.bar(r.iter_content(chunk_size=2391975),
                                           expected_size=(total_length / 1024) + 1):
                        if ch:
                            Pypdf.write(ch)

            # NPPR9-FWDCX-D2C8J-H872K-2YT43
            elif "jarvis" in self.query:

                wishMe()
                speak("Jarvis 1 point o in your service Mister")
                speak(assname)

            elif "weather" in self.query:

                # Google Open weather website
                # to get API of Open weather
                api_key = "27e77758bd211846998f3629b6c519af"
                base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
                speak(" City name ")
                city_name = self.takeCommand()
                complete_url = base_url + "appid =" + api_key + "&q =" + city_name
                response = requests.get(complete_url)
                x = response.json()

                if x["cod"] != "404":
                    y = x["main"]
                    current_temperature = y["temp"]
                    current_pressure = y["pressure"]
                    current_humidiy = y["humidity"]
                    z = x["weather"]
                    weather_description = z[0]["description"]
                    print(" Temperature (in kelvin unit) = " + str(
                        current_temperature) + "\n atmospheric pressure (in hPa unit) =" + str(
                        current_pressure) + "\n humidity (in percentage) = " + str(
                        current_humidiy) + "\n description = " + str(weather_description))

                else:
                    speak(" City Not Found ")

            elif "send message" in self.query:
                # You need to create an account on Twilio to use this service
                account_sid = os.environ['AC765eb606691c6a0d3638ddff5426843e']
                auth_token = os.environ['c7179a2091973228050a14907679b8af']
                client = Client(account_sid, auth_token)

                message = client.messages \
                    .create(
                        body='This is the testing',
                        from_='+13017615956',
                        to='+919008059668'
                    )

                print(message.sid)

            elif "wikipedia" in self.query:
                webbrowser.open("wikipedia.com")

            elif "Good Morning" in self.query:
                speak("A warm" + query)
                speak("How are you Mister")
                speak(assname)

            # most asked question from google Assistant
            elif "will you be my gf" in self.query or "will you be my bf" in self.query:
                speak("I'm not sure about, may be you should give me some time")

            elif "how are you" in self.query:
                speak("I'm fine, glad you me that")

            elif "i love you" in self.query:
                speak("It's hard to understand")

            elif "what is" in self.query or "who is" in self.query:

                # Use the same API key
                # that we have generated earlier
                client = wolframalpha.Client("EPHWP8-W46635H8VG")
                res = client.query(self.query)

                try:
                    speak(next(res.results).text)
                except StopIteration:
                    print("No results")

            elif "ip address" in self.query:
                ip = get('https://api.ipify.org').text
                speak(f'Your ip address is {ip}')

            elif "open notepad" in self.query:
                speak("Ok sir, i got it")
                os.startfile("C:\\Windows\\system32\\notepad.exe")

            elif "close notepad" in self.query:
                speak("Ok sir, i got it")
                os.system("taskkill /f /im notepad.exe")

            elif "close browser" in self.query:
                speak("Ok sir, i got it")
                os.system("taskkill /f /im chrome.exe")

            elif "open adobe reader" in self.query:
                os.startfile("C:\\Program Files (x86)\\Adobe\\Acrobat DC\\Acrobat\\Acrobat.exe")

            elif "close adobe reader" in self.query:
                speak("Ok sir, i got it")
                os.system("taskkill /f /im Acrobat.exe")

            elif "open command prompt" in self.query:
                os.system("start cmd")

            elif "open instagram" in self.query:
                webbrowser.open("instagram.com")

            elif "open facebook" in self.query:
                webbrowser.open("facebook.com")

            elif "open google" in self.query:
                speak("Sir what should i search on google")
                cm = self.takeCommand().lower()
                webbrowser.open(f"{cm}")

            elif "play on youtube" in self.query:
                speak("Sir which song you want to listen.")
                u = self.takeCommand().lower()
                kit.playonyt(f"{u}")

            elif "no thanks" in self.query:
                speak("Thanks for assisting me sir, have a good day.")
                sys.exit()

            elif "where i am" in self.query or "where we are" in self.query:
                speak("wait sir, let me check")
                try:
                    ipAdd = requests.get('https://api.ipify.org').text
                    print(ipAdd)
                    url = 'https://get.geojs.io/v1/ip/geo/' + ipAdd + '.json'
                    geo_requests = requests.get(url)
                    geo_data = geo_requests.json()
                    city = geo_data['city']
                    country = geo_data['country']
                    speak(f"Sir im not sure but i think we are in {city} city of {country} country")

                except Exception as e:
                    speak("Sorry sir, Due to network issue i am unable to find where we are.")
                    pass

            elif "instagram profile" in self.query or "profile on instagram" in self.query:
                speak("Please enter username correctly")
                name = input("Enter user name:")
                webbrowser.open(f"https://www.instagram.com/{name}")
                speak(f"Sir here is the profile of the user {name}")
                time.sleep(5)
                speak("sir would you like to download the profile picture of the account?")
                condition = self.takeCommand().lower()
                if "yes" in condition:
                    mod = instaloader.Instaloader()
                    mod.download_profile(name, profile_pic_only=True)
                    speak("its done sir, saved picture can locate in main folder. Now im ready")
                else:
                    pass

            elif "take screenshot" in self.query or "take a screenshot" in self.query:
                speak("Sir please tell me the name of the screenshot file")
                name = self.takeCommand().lower()
                speak("Sir please hold the screen for few seconds, i am taking screenshot")
                time.sleep(3)
                img = pyautogui.screenshot()
                img.save(f"{name}.png")
                speak("I am done sir, screenshot is saved in main folder, im ready for next command")

            elif "switch the window" in self.query:
                pyautogui.keyDown("alt")
                pyautogui.press("tab")
                time.sleep(1)
                pyautogui.keyUp("alt")

            elif "what\'s up" in self.query or 'how are you' in self.query:
                stMsgs = ['Just doing my thing!', 'I am fine!', 'Nice!', 'I am nice and full of energy']
                speak(random.choice(stMsgs))

            elif 'nothing' in self.query or 'abort' in self.query or 'stop' in self.query:
                speak('okay')
                speak('Bye Boss, have a good day.')
                sys.exit()

            elif 'hello' in self.query:
                speak('Hello Boss')

            elif 'bye' in self.query:
                speak('Bye Boss, have a good day.')
                sys.exit()

            elif 'voice' in self.query:
                if 'female' in self.query:
                    engine.setProperty('voice', voices[1].id)
                else:
                    engine.setProperty('voice', voices[0].id)
                speak("Hello Sir, I have switched my voice. How is it?")


            elif 'remember that' in self.query:
                speak("what should i remember sir")
                rememberMessage = self.takeCommand()
                speak("you said me to remember" + rememberMessage)
                remember = open('data.txt', 'w')
                remember.write(rememberMessage)
                remember.close()

            elif 'do you remember anything' in self.query:
                remember = open('data.txt', 'r')
                speak("you said me to remember that" + remember.read())

startExecution = MainThread()

class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.startTask)
        self.ui.pushButton_2.clicked.connect(self.close)

    def startTask(self):
        self.ui.movie = QtGui.QMovie("7LP8.gif")
        self.ui.label.setMovie(self.ui.movie)
        self.ui.movie.start()
        timer = QTimer(self)
        timer.timeout.connect(self.showTime)
        timer.start(1000)
        startExecution.start()

    def showTime(self):
        current_time = QTime.currentTime()
        current_date = QDate.currentDate()
        lable_time = current_time.toString('hh:mm:ss')
        lable_date = current_date.toString(Qt.ISODate)
        self.ui.textBrowser.setText(lable_date)
        self.ui.textBrowser_2.setText(lable_time)


app = QApplication(sys.argv)
jarvis = Main()
jarvis.show()
exit(app.exec_())




    
