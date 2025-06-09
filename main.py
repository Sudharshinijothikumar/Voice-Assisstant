import subprocess
import pyautogui
import wolframalpha
import pyttsx3
import random
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import json
import feedparser
import smtplib
import datetime 
import requests
import shutil
from twilio.rest import Client
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import ecapture as ec
import screen_brightness_control as sbc
import time
import getpass
import re  

# Initialize voice engine
voiceEngine = pyttsx3.init('sapi5')
voices = voiceEngine.getProperty('voices')
voiceEngine.setProperty('voice', voices[0].id)

def speak(text):
    voiceEngine.say(text)
    voiceEngine.runAndWait()

def wish():
    print("Hello.")
    time = int(datetime.datetime.now().hour)
    global uname, asname
    if time >= 0 and time < 12:
        speak("Good Morning sir or madam!")
    elif time < 18:
        speak("Good Afternoon sir or madam!")
    else:
        speak("Good Evening sir or madam!")

    asname = "Laddoo"
    speak("I am your")
    speak(asname)
    print("I am your Voice Assistant,", asname)

def getName():
    global uname
    speak("Can I please know your name?")
    uname = takeCommand()
    print("Name:", uname)
    speak("I am glad to know you!")
    speak("How can I help you, ")
    speak(uname)

def process_command(command):
    # Normalize whitespace first
    command = ' '.join(str(command).split())
    
    # Handle special cases where symbol words might be attached to other words
    patterns = [
        (r'(\w)at(\w)', r'\1@\2'),    # useratdomain -> user@domain
        (r'(\w)dot(\w)', r'\1.\2'),    # domaindotcom -> domain.com
        (r'(\w)slash(\w)', r'\1/\2'),  # pathslashfile -> path/file
        (r'(\w)hash(\w)', r'\1#\2')    # pythonhashset -> python#set
    ]
    
    for pattern, replacement in patterns:
        command = re.sub(pattern, replacement, command)
    
    # Standard replacements
    replacements = {
        " at ": "@",
        " dot ": ".",
        " slash ": "/",
        " hash ": "#",
        " hashtag ": "#",
        " percent ": "%",
        " underscore ": "_",
        " dash ": "-",
        " plus ": "+",
        " equals ": "=",
        " colon ": ":",
        " semicolon ": ";",
        " space ": " "
    }
    
    for k, v in replacements.items():
        command = command.replace(k, v)
    
    return command.lower()

def takeCommand():
    recog = sr.Recognizer()
    
    with sr.Microphone() as source:
        print("Listening to the user...")
        recog.pause_threshold = 1
        recog.adjust_for_ambient_noise(source)
        userInput = recog.listen(source)

    try:
        print("Recognizing the command...")
        command = recog.recognize_google(userInput, language='en-in')
        command = process_command(command)
        print(f"Command is: {command}\n")
    except Exception as e:
        print(e)
        print("Unable to Recognize the voice.")
        return "None"

    return command

def sendEmail(to, content):
    try:
        print("Preparing to send email...")
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        
        # Ask user to speak their email address
        speak("Please say your email address")
        email = confirm_email_address()
        
        # Ask user to speak their email password
        speak("Please say your email password, character by character. Say 'done' when finished.")
        password_chars = []
        while True:
            char = takeCommand()
            if char.lower() == "done":
                break
            # Remove spaces and take only first character if user says a word
            char = char.replace(" ", "")
            if len(char) > 1:
                char = char[0]
            password_chars.append(char)
        password = ''.join(password_chars)
        
        server.login(email, password)
        server.sendmail(email, to, content)
        server.close()
        print("Email sent successfully!")
        return True
    except Exception as e:
        print(f"Email error: {str(e)}")
        return False

def getWeather(city_name):
    baseUrl = "http://api.openweathermap.org/data/2.5/weather?"
    url = baseUrl + "appid=" + 'd850f7f52bf19300a9eb4b0aa6b80f0d' + "&q=" + city_name  
    response = requests.get(url)
    x = response.json()

    if x["cod"] != "404":
        y = x["main"]
        temp = y["temp"] - 273  # Convert Kelvin to Celsius
        pressure = y["pressure"]
        humidity = y["humidity"]
        desc = x["weather"][0]["description"]
        info = (f"Temperature: {temp:.1f}°C\nPressure: {pressure} hPa\nHumidity: {humidity}%\nDescription: {desc}")
        print(info)
        speak(f"Here is the weather report for {city_name}")
        speak(info)
    else:
        speak("City Not Found")

def getNews():
    try:
        response = requests.get('https://www.bbc.com/news')
        soup = BeautifulSoup(response.text, 'html.parser')
        headlines = soup.find('body').find_all('h3')
        unwanted = ['BBC World News TV', 'BBC World Service Radio',
                   'News daily newsletter', 'Mobile app', 'Get in touch']

        for x in list(dict.fromkeys(headlines)):
            if x.text.strip() not in unwanted:
                print(x.text.strip())
                speak(x.text.strip())
    except Exception as e:
        print(str(e))
        speak("Sorry, I couldn't fetch the news right now.")

def take_screenshot():
    try:
        if not os.path.exists('screenshots'):
            os.makedirs('screenshots')
        
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"screenshots/screenshot_{timestamp}.png"
        
        pyautogui.screenshot(filename)
        speak(f"Screenshot saved as {filename}")
        return True
    except Exception as e:
        print(f"Screenshot error: {e}")
        speak("Sorry, I couldn't take a screenshot.")
        return False

def wolfram_query(query):
    try:
        client = wolframalpha.Client('YOUR_WOLFRAM_APP_ID')
        res = client.query(query)
        
        if not res['@success']:
            return "I couldn't find an answer for that."
            
        answer = next(res.results).text
        return answer
    except Exception as e:
        print(f"Wolfram Alpha error: {e}")
        return "Sorry, I couldn't process that request."

def confirm_email_address():
    speak("Please say the email address")
    while True:
        email = takeCommand()
        # Remove all whitespace and normalize
        processed_email = re.sub(r'\s+', '', email).replace('at', '@').replace('dot', '.')
        
        # Basic email validation
        if re.match(r'[^@]+@[^@]+\.[^@]+', processed_email):
            speak(f"I heard {processed_email}. Is this correct? Please say yes or no.")
            confirmation = takeCommand().lower()
            
            if "yes" in confirmation:
                return processed_email
            elif "no" in confirmation:
                speak("Please say the email address again")
            else:
                speak("I didn't understand. Please say yes or no.")
        else:
            speak("That doesn't appear to be a valid email address. Please try again.")

if __name__ == '__main__':
    uname = ''
    asname = ''
    os.system('cls')
    wish()
    getName()

    while True:
        command = takeCommand()
        print("User said:", command)

        if not command or command == "none":
            continue
            
        if "jarvis" in command:
            wish()
            
        elif 'how are you' in command:
            speak("I am fine, Thank you")
            speak(f"How are you, {uname}")

        elif any(greeting in command for greeting in ["good morning", "good afternoon", "good evening"]):
            speak(f"A very {command}")
            speak("Thank you for wishing me! Hope you are doing well!")

        elif 'fine' in command or "good" in command:
            speak("It's good to know that you're fine")
       
        elif "who are you" in command:
            speak("I am your virtual assistant.")

        elif "change my name to" in command:
            speak("What would you like me to call you?")
            uname = takeCommand()
            speak(f'Hello again, {uname}')
        
        elif "change name" in command:
            speak("What would you like to call me?")
            asname = takeCommand()
            speak("Thank you for naming me!")

        elif "what's your name" in command:
            speak(f"People call me {asname}")
        
        elif 'time' in command:
            current_time = datetime.datetime.now().strftime("%I:%M %p")
            speak(f"{uname}, the time is {current_time}")
            print(current_time)

        elif 'wikipedia' in command:
            speak('Searching Wikipedia...')
            query = command.replace("wikipedia", "")
            try:
                results = wikipedia.summary(query, sentences=3)
                speak("According to Wikipedia")
                print(results)
                speak(results)
            except Exception as e:
                print(e)
                speak("Sorry, I couldn't find that on Wikipedia")

        elif 'open youtube' in command:
            speak("Opening YouTube")
            webbrowser.open("youtube.com")

        elif 'open google' in command:
            speak("Opening Google")
            webbrowser.open("google.com")

        elif any(cmd in command for cmd in ['play music', 'play song','play some music', 'play a song', 'play a tune', 'play a track']):
            speak("Playing music")
            music_dir = "C:\\Users\\Public\\Music"
            songs = os.listdir(music_dir)
            if songs:
                random_song = random.choice(songs)
                os.startfile(os.path.join(music_dir, random_song))
            else:
                speak("No songs found in your music directory")

        elif 'joke' in command:
            speak(pyjokes.get_joke())
            
        elif any(cmd in command for cmd in ['mail', 'email']):
            try:
                to = confirm_email_address()
                speak("What should I say in the email?")
                content = takeCommand()
                
                if sendEmail(to, content):
                    speak("Email sent successfully!")
                else:
                    speak("Failed to send the email")
                    
            except Exception as e:
                print("Email error:", e)
                speak("Sorry, I encountered an error while sending the email")

        elif any(cmd in command for cmd in ['exit', 'quit', 'stop', 'close', 'bye','tata', 'goodbye', 'see you later']):
            speak("Thanks for your time. Have a great day!")
            exit()

        elif "cloud" in command or "sky" in command or "cloudy" in command or "weather" in command or "forecast" in command or "climate" in command or "temperature" in command: 
            speak("Please tell me the city name")
            cityName = takeCommand()
            getWeather(cityName)

        elif any(prefix in command for prefix in ['calculate', 'what is', 'who is', 'how many', 'convert']):
            query = command.replace('calculate', '').strip()
            result = wolfram_query(query)
            speak(result)
            print(result)

        elif 'search' in command:
            query = command.replace("search", "").strip()
            if query:
                speak(f"Searching for {query}")
                webbrowser.open(f"https://www.google.com/search?q={query}")
            else:
                speak("What would you like me to search for?")

        elif 'news' in command:
            getNews()
        
        elif any(cmd in command for cmd in ["don't listen", "stop listening"]):
            speak("For how many seconds should I stop listening?")
            try:
                sleep_time = int(takeCommand())
                time.sleep(sleep_time)
            except:
                speak("I didn't understand that time period")

        elif any(cmd in command for cmd in ["camera", "take a photo"]):
            ec.capture(0, "Jarvis Camera ", "img.jpg")
        
        elif 'shutdown system' in command:
            speak("Shutting down the system")
            os.system("shutdown /s /t 1")

        elif "restart" in command:
            speak("Restarting the system")
            os.system("shutdown /r /t 1")

        elif "sleep" in command:
            speak("Putting system to sleep")
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

        elif "open" in command:
            speak("What would you like me to open?")
            app = takeCommand().lower()
            if "chrome" in app:
                os.startfile("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe")
            elif "notepad" in app:  
                os.startfile("notepad.exe")           
            elif "word" in app:
                os.startfile("winword.exe")
            elif "write a note" in command:
                speak("What should I write?")
                note = takeCommand()
                with open('jarvis.txt', 'w') as file:
                    speak("Should I include date and time?")
                    if 'yes' in takeCommand().lower():
                        file.write(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}: {note}")
                    else:
                        file.write(note)
                speak("Note saved successfully")
            
        elif any(cmd in command for cmd in ['open settings', 'system settings']):
            speak("Opening Windows Settings")
            os.system("start ms-settings:")

        elif 'lock pc' in command:
            speak("Locking your computer")
            os.system("rundll32.exe user32.dll,LockWorkStation")

        elif 'increase brightness' in command or 'increase the brightness' in command or 'increase screen brightness' in command or 'increase display brightness' in command or 'increase monitor brightness' in command:
            current = sbc.get_brightness()[0]
            new = min(current + 20, 100)
            sbc.set_brightness(new)
            speak(f"Brightness increased to {new}%")

        elif 'decrease brightness' in command or 'decrease the brightness' in command or 'decrease screen brightness' in command or 'decrease display brightness' in command or 'decrease monitor brightness' in command:
            current = sbc.get_brightness()[0]
            new = max(current - 20, 0)
            sbc.set_brightness(new)
            speak(f"Brightness decreased to {new}%")

        elif 'increase volume' in command or 'volume up' in command or 'louder' in command or 'increase sound' in command or 'increase audio' in command:
            pyautogui.press('volumeup')
            speak("Volume increased")

        elif 'decrease volume' in command or 'volume down' in command or 'quieter' in command or 'decrease sound' in command or 'decrease audio' in command:
            pyautogui.press('volumedown')
            speak("Volume decreased")

        elif 'mute' in command or 'unmute' in command:
            pyautogui.press('volumemute')
            speak("Volume toggled") 

        elif any(cmd in command for cmd in ['open google calendar', 'open calendar']):
            speak("Opening Google Calendar")
            webbrowser.open("https://calendar.google.com")

        elif any(cmd in command for cmd in ['open vs code', 'open visual studio code']):
            speak("Opening VS Code")
            os.system("code")

        elif any(cmd in command for cmd in ['open google drive', 'open drive']):
            speak("Opening Google Drive")
            webbrowser.open("https://drive.google.com")

        elif 'open calculator' in command:
            speak("Opening Calculator")
            os.system("calc")

        elif 'open notepad' in command:
            speak("Opening Notepad")
            os.system("notepad")

        elif 'open file explorer' in command:
            speak("Opening File Explorer")
            os.system("explorer")

        elif any(cmd in command for cmd in ['take screenshot', 'capture screen']):
            if take_screenshot():
                speak("Screenshot captured successfully")
            else:
                speak("Failed to capture screenshot")

        else:
            speak("I didn't understand that. Could you please repeat?")
