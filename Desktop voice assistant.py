import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import pyautogui
import smtplib
import requests
import psutil
import pyjokes
import tkinter as tk
from tkinter import messagebox
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import ctypes
import pygetwindow as gw
import random
from googletrans import Translator
import openai

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

# OpenAI API key
openai.api_key = 'your_openai_api_key'

def speak(text):
    engine.say(text)
    engine.runAndWait()

def wishMe():
    hour = datetime.datetime.now().hour
    if 0 <= hour < 12:
        speak("Good Morning!")
    elif 12 <= hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak("I am Jarvis, Sir. Please tell me how may I help you")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")
    except Exception as e:
        print(e)
        print("Say that again please...")
        return "None"
    return query.lower()

def sendEmail(to, content):
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        server.login('your_email@gmail.com', 'your_password')
        server.sendmail('your_email@gmail.com', to, content)
        server.close()
        speak("Email has been sent!")
    except Exception as e:
        print(e)
        speak("Sorry sir, I am not able to send this email at the moment")

def showMessageBox(message):
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Message from Jarvis", message)
    root.destroy()

def getNewsHeadlines():
    url = "https://newsapi.org/v2/top-headlines?country=us&apiKey=your_newsapi_key"
    response = requests.get(url)
    news_data = response.json()
    headlines = [article['title'] for article in news_data['articles']]
    return headlines

def confirmAction(action):
    speak(f"Are you sure you want to {action}?")
    confirmation = takeCommand().lower()
    return 'yes' in confirmation

def changeVolume(level):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        volume.SetMasterVolume(level, None)

def getVolume():
    sessions = AudioUtilities.GetAllSessions()
    volume = sessions[0]._ctl.QueryInterface(ISimpleAudioVolume)
    return volume.GetMasterVolume()

def setBrightness(level):
    if 0 <= level <= 100:
        ctypes.windll.graphics.SetDeviceGammaRamp(0, ctypes.c_void_p(), ctypes.c_void_p(int(level * 655.35)))
        speak(f"Screen brightness set to {level}%")
    else:
        speak("Invalid brightness level. Please provide a value between 0 and 100.")

def setResolution(width, height):
    screen = gw.getWindowsWithTitle('')[0]
    screen.moveTo(0, 0)
    screen.resizeTo(width, height)
    speak(f"Screen resolution set to {width} by {height} pixels")

def tellDate():
    now = datetime.datetime.now()
    date = now.strftime("%A, %B %d, %Y")
    speak(f"Today's date is {date}")

reminders = {}

def addReminder():
    speak("Sure, what should I remind you about?")
    reminder_text = takeCommand()
    speak("When should I remind you?")
    reminder_time = takeCommand()
    try:
        reminder_time = datetime.datetime.strptime(reminder_time, "%H:%M")
        reminders[reminder_time] = reminder_text
        speak("Reminder set!")
    except Exception as e:
        print(e)
        speak("Sorry, I couldn't understand the time format. Please provide the time in HH:MM format.")

def checkReminders():
    current_time = datetime.datetime.now().time()
    for reminder_time, reminder_text in list(reminders.items()):
        if current_time >= reminder_time.time():
            speak(f"Reminder: {reminder_text}")
            del reminders[reminder_time]

def playGame():
    speak("Let's play a game! I'll choose a number between 1 and 10. You guess it.")
    number = random.randint(1, 10)
    guesses = 0
    while True:
        speak("Take a guess.")
        guess = takeCommand()
        if not guess.isdigit():
            speak("Please provide a number.")
            continue
        guess = int(guess)
        guesses += 1
        if guess < number:
            speak("Too low.")
        elif guess > number:
            speak("Too high.")
        else:
            speak(f"Good job! You guessed the number in {guesses} guesses.")
            break

def tellJoke():
    speak(pyjokes.get_joke())

def controlHomeAutomation():
    speak("Which appliance would you like to control?")
    appliance = takeCommand().lower()
    if 'lights' in appliance:
        speak("Turning lights on/off.")
    elif 'ac' in appliance or 'air conditioner' in appliance:
        speak("Turning air conditioner on/off.")
    elif 'television' in appliance or 'tv' in appliance:
        speak("Turning television on/off.")
    else:
        speak("Sorry, I don't know how to control that appliance yet.")

def translateText(text, dest_language):
    translator = Translator()
    translated_text = translator.translate(text, dest=dest_language)
    return translated_text.text

def suggestWorkout():
    speak("Let's do a quick workout! How about some push-ups?")

def suggestHealthTips():
    speak("Here's a health tip: Drink plenty of water throughout the day.")

def suggestHealthyRecipe():
    speak("How about trying a salad for a healthy meal?")

def suggestActivity():
    speak("Here are some activities you might enjoy.")
    activities = ["reading a book", "going for a walk", "listening to music", "practicing meditation"]
    suggested_activity = random.choice(activities)
    speak(f"How about {suggested_activity}?")

def suggestHealthyHabit():
    speak("Here's a healthy habit you can try.")
    habits = ["drinking a glass of water", "doing some stretches", "taking a deep breath"]
    suggested_habit = random.choice(habits)
    speak(f"Why not try {suggested_habit}?")

def suggestNutritiousFood():
    speak("How about trying a nutritious meal?")
    meals = ["a salad with lots of vegetables", "a smoothie with fruits and yogurt", "grilled chicken with vegetables"]
    suggested_meal = random.choice(meals)
    speak(f"You might like {suggested_meal}.")

def generate_response(prompt):
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        max_tokens=150
    )

def automateTasks():
    speak("Sure, I can help with that. What task would you like to automate?")
    task = takeCommand().lower()
    if 'shutdown' in task:
        if confirmAction("shutdown"):
            os.system("shutdown /s /t 1")
            speak("Shutting down the computer.")
    elif 'restart' in task:
        if confirmAction("restart"):
            os.system("shutdown /r /t 1")
            speak("Restarting the computer.")
    elif 'log off' in task:
        if confirmAction("log off"):
            os.system("shutdown /l")
            speak("Logging off.")
    elif 'hibernate' in task:
        if confirmAction("hibernate"):
            os.system("shutdown /h")
            speak("Hibernating the computer.")
    else:
        speak("Sorry, I don't know how to automate that task yet.")

if __name__ == "__main__":
    wishMe()
    while True:
        query = takeCommand()
        if 'add reminder' in query:
            addReminder()
        
        elif 'check reminders' in query:
            checkReminders()
        
        elif 'play game' in query:
            playGame()
        
        elif 'tell me a joke' in query:
            tellJoke()

        elif 'control' in query and 'home' in query:
            controlHomeAutomation()
        
        elif'translate' in query:
            speak("What should I translate?")
            text_to_translate = takeCommand()
            speak("Which language should I translate it to?")
            dest_language = takeCommand()
            translated_text = translateText(text_to_translate, dest_language)
            speak(f"The translation is: {translated_text}")

        elif 'suggest workout' in query:
            suggestWorkout()

        elif 'suggest health tips' in query:
            suggestHealthTips()

        elif 'suggest healthy recipe' in query:
            suggestHealthyRecipe()    

        elif 'suggest activity' in query:
            suggestActivity()

        elif 'suggest healthy habit' in query:
            suggestHealthyHabit()

        elif 'suggest nutritious food' in query:
            suggestNutritiousFood()

        
        elif 'task automation' in query:
            automateTasks()

        elif 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak("Opening YouTube, sir....")
            webbrowser.open("https://www.youtube.com")

        elif 'open google' in query:
            speak("Opening Google, sir....")
            webbrowser.open("https://www.google.com")

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        elif 'open code' in query:
            speak("Opening Visual Studio Code, sir....")
            codePath = "C:\\Users\\DEll\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)
        
        elif 'play music' in query:
         speak("Playing music, sir....")
         music_dir = 'E:\\my music'
         songs = os.listdir(music_dir)
         for song in songs:
          os.startfile(os.path.join(music_dir, song))

        elif 'move mouse' in query:
            speak("Moving the mouse to the center of the screen, sir....")
            pyautogui.moveTo(pyautogui.size()[0] // 2, pyautogui.size()[1] // 2)

        elif 'click' in query:
            speak("Performing a mouse click, sir....")
            pyautogui.click()

        elif 'screenshot' in query:
            speak("Taking a screenshot, sir....")
            screenshot_path = "C:\\Users\\DEll\\Desktop\\screenshot.png"
            pyautogui.screenshot(screenshot_path)
            speak("Screenshot saved on the desktop")

        elif 'battery' in query:
            battery = psutil.sensors_battery()
            percentage = battery.percent
            speak(f"Sir, our current battery percentage is {percentage}")

        elif 'email' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "abhayshukla994@gmail.com"
                sendEmail(to, content)
                speak("Email has been sent!")

            except Exception as e:
                print(e)
                speak("Sorry sir, I am not able to send this email at the moment")
        
        elif 'weather' in query:
           speak("Checking the weather, sir....")
           city = "Lucknow"  
           API_Key = "79f9e09c776385576e3cf239708b82ed"
           url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={API_Key}&units=metric"

           response = requests.get(url)
           weather_data = response.json()
           temperature = weather_data['main']['temp']
           weather_desc = weather_data['weather'][0]['description']
           speak(f"The temperature in {city} is {temperature} degrees Celsius with {weather_desc}")
           

        elif 'search' in query:
            speak("What should I search for?")
            search_query = takeCommand()
            webbrowser.open(f"https://www.google.com/search?q={search_query}")

        elif 'open website' in query:
            speak("Which website should I open?")
            website = takeCommand().replace(" ", "")
            webbrowser.open(f"https://{website}.com")

        elif 'open notepad' in query:
            speak("Opening Notepad, sir....")
            os.system("notepad.exe")

        elif 'open calculator' in query:
            speak("Opening Calculator, sir....")
            os.system("calc.exe")

        elif 'open browser' in query:
            speak("Opening Browser, sir....")
            os.system("start chrome")

        elif 'open word' in query:
            speak("Opening Microsoft Word, sir....")
            os.system("start winword")

        elif 'open excel' in query:
            speak("Opening Microsoft Excel, sir....")
            os.system("start excel")

        elif 'open powerpoint' in query:
            speak("Opening Microsoft PowerPoint, sir....")
            os.system("start powerpnt")

        elif 'move left' in query:
            speak("Moving the mouse cursor to the left, sir....")
            current_x, current_y = pyautogui.position()
            pyautogui.moveTo(current_x - 50, current_y, duration=0.25)

        elif 'move right' in query:
            speak("Moving the mouse cursor to the right, sir....")
            current_x, current_y = pyautogui.position()
            pyautogui.moveTo(current_x + 50, current_y, duration=0.25)

        elif 'move up' in query:
            speak("Moving the mouse cursor up, sir....")
            current_x, current_y = pyautogui.position()
            pyautogui.moveTo(current_x, current_y - 50, duration=0.25)

        elif 'move down' in query:
            speak("Moving the mouse cursor down, sir....")
            current_x, current_y = pyautogui.position()
            pyautogui.moveTo(current_x, current_y + 50, duration=0.25)

        elif 'scroll up' in query:
            speak("Scrolling up, sir....")
            pyautogui.scroll(100)

        elif 'scroll down' in query:
            speak("Scrolling down, sir....")
            pyautogui.scroll(-10)

        elif 'show message' in query:
            speak("What message should I display?")
            message = takeCommand()
            showMessageBox(message)

        elif 'news' in query:
            headlines = getNewsHeadlines()
            for headline in headlines:
                speak(headline)
                
        elif 'change volume' in query:
            speak("Sure, what should be the new volume level?")
            volume_level = float(takeCommand())
            if 0.0 <= volume_level <= 1.0:
                changeVolume(volume_level)
                speak(f"Volume set to {volume_level}")
            else:
                speak("Invalid volume level. Please provide a value between 0.0 and 1.0")

        elif 'current volume' in query:
            current_volume = getVolume()
            speak(f"Sir, the current volume level is {current_volume}")

        elif 'change brightness' in query:
            speak("Sure, what should be the new brightness level?")
            brightness_level = int(takeCommand())
            setBrightness(brightness_level)

        elif 'change resolution' in query:
            speak("Sure, what should be the new screen resolution? Please provide width and height in pixels.")
            try:
                width, height = map(int, takeCommand().split())
                setResolution(width, height)
            except Exception as e:
                print(e)
                speak("Invalid input. Please provide the width and height as two numbers separated by space.")
     
        elif 'tell me the date' in query:
            tellDate()
  
        elif 'thank you' in query:
            speak("Thank you, sir. Have a great day!")
            break

        elif 'exit' in query or 'quit' in query:
            speak("Goodbye, Sir.")
            break

        else:
            response = generate_response(query)
            speak(response)
            showMessageBox(response)
