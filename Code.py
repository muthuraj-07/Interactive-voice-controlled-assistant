import datetime
import os
import time
import pyttsx3
import speech_recognition as sr
import wikipedia
import webbrowser
import pywhatkit as wk
import sys
import random
import pyautogui
import requests
import cv2

# Initialize the text-to-speech engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)  # Use the first voice available
engine.setProperty('rate', 150)  # Set speaking rate

def speak(audio):
    """Convert text to speech."""
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak('Good Morning boss')
    elif hour >= 12 and hour < 18:
        speak('Good afternoon boss')
    else:
        speak('Good evening boss')
    speak("what can I do for you ?")

def takeCommand():
    """Listen to the microphone and recognize speech."""
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)  # Capture audio

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')  # Recognize speech
        print(f"User said: {query}\n")  # Print the recognized speech
    except sr.UnknownValueError:
        print("Google Speech Recognition could not understand the audio")
        return "None"
    except sr.RequestError:
        print("Could not request results; check your network connection")
        return "None"
    except Exception as e:
        print(f"An error occurred: {e}")
        return "None"

    return query

if __name__ == "__main__":
    wishMe()
    while True:
        query = takeCommand().lower()

        # Basic intro of Nova (a virtual voice assistant)
        if 'hello nova' in query:
            print("Yes sir, how can I help you")
            speak("Yes sir, how can I help you")
        elif 'who are you' in query:
            print('My name is Nova')
            speak('My name is Nova')
            print('I can perform tasks that my creator programmed me to do')
            speak('I can perform tasks that my creator programmed me to do')
        elif 'who created you' in query:
            print('I was created by a group of BITM students')
            speak('I was created by a group of BITM students')

        elif 'who is' in query:
            speak('Searching Wikipedia...')
            query = query.replace("who is", "").strip()
            try:
                results = wikipedia.summary(query, sentences=2)
                speak("According to Wikipedia")
                print(results)
                speak(results)
            except wikipedia.exceptions.PageError:
                speak("Sorry, I couldn't find any information on that topic.")
            except wikipedia.exceptions.DisambiguationError as e:
                speak("The query is ambiguous. Please be more specific.")
                print(f"DisambiguationError: {e.options}")
            except Exception as e:
                speak("An error occurred.")
                print(f"An error occurred: {e}")

        elif 'just open google' in query:
            speak('Opening Google')
            webbrowser.open('google.com')

        elif 'open google' in query:
            speak('What should I search for?')
            qry = takeCommand().lower()
            search_url = f"https://www.google.com/search?q={qry}"
            webbrowser.open(search_url)
            try:
                results = wikipedia.summary(qry, sentences=1)
                speak(results)
            except wikipedia.exceptions.DisambiguationError as e:
                speak(f"Your query is too vague. Please be more specific. Suggestions: {e.options}")
            except wikipedia.exceptions.PageError:
                speak("Sorry, I couldn't find any results on Wikipedia.")
            except Exception as e:
                speak("An error occurred.")

        elif 'just open youtube' in query:
            speak("Opening YouTube")
            webbrowser.open("https://www.youtube.com")
        elif 'open youtube' in query:
            speak("What would you like to watch?")
            qrry = takeCommand().lower()
            wk.playonyt(f"{qrry}")

        elif 'search on youtube' in query:
            query = query.replace('search on youtube', "")
            webbrowser.open(f"www.youtube.com/results?search_query={query}")

        elif 'what the time now' in query:
            now = datetime.datetime.now()
            current_time = now.strftime("%H:%M")
            speak(f"The time is {current_time}")

        elif 'open excel' in query:
            try:
                os.startfile(r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE")
                speak("Excel started successfully.")
            except Exception as e:
                speak(f"An error occurred while opening Excel: {e}")
        elif 'close excel' in query:
            try:
                os.system("taskkill /f /im EXCEL.EXE")
                speak("Excel closed successfully.")
            except Exception as e:
                speak(f"An error occurred while closing Excel: {e}")

        elif 'open command prompt' in query:
            os.startfile(r"C:\windows\system32\cmd.exe")
            speak("Command prompt started successfully")
        elif 'close command prompt' in query:
            os.system("taskkill /f /im cmd.exe")
            speak("Command prompt has been closed")

        elif 'open powerpoint' in query:
            os.startfile(r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE")
            speak("PowerPoint started successfully")
        elif 'close powerpoint' in query:
            os.system("taskkill /f /im POWERPNT.EXE")
            speak("PowerPoint has been closed")

        elif 'open vs code' in query:
            os.startfile(r"C:\Users\manoh\AppData\Local\Programs\Microsoft VS Code\Code.exe")
            speak("VS Code started successfully")
        elif 'close vs code' in query:
            os.system("taskkill /f /im Code.exe")
            speak("VS Code has been closed")

        elif 'open camera' in query:
            cap = cv2.VideoCapture(0)
            if not cap.isOpened():
                print("Error: Could not open video capture.")
                exit()
            while True:
                ret, img = cap.read()
                if not ret:
                    print("Error: Could not read frame.")
                    break
                cv2.imshow('Webcam', img)
                k = cv2.waitKey(1)
                if k == 27:
                    break
            cap.release()
            cv2.destroyAllWindows()

        elif 'what is my ip address' in query:
            speak('Checking boss')
            try:
                ipAdd = requests.get('https://api.ipify.org').text
                print(ipAdd)
                speak('Your IP address is:')
                speak(ipAdd)
            except Exception as e:
                speak('Network is slow, please try again later')

        elif 'take screenshot' in query:
            speak('Tell me a name for the file')
            name = takeCommand().lower()
            time.sleep(3)
            img = pyautogui.screenshot()
            img.save(f"C:/Users/manoh/Desktop/{name}.png")
            speak('Screenshot saved')

        elif 'volume up' in query:
            for _ in range(10):
                pyautogui.press('volumeup')

        elif 'volume down' in query:
            for _ in range(10):
                pyautogui.press('volumedown')

        elif 'mute' in query:
            pyautogui.press('volumemute')

        elif 'ok you can exit' in query:
            speak("OK boss, see you later")
            break
