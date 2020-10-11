import pyttsx3 #pip install pyttsx3
import speech_recognition as sr #pip install speechRecognition
import datetime
import wikipedia #pip install wikipedia
import webbrowser
import os
import smtplib
import subprocess
import shutil
import pyjokes
import ctypes
from urllib.request import urlopen
import json
import requests
import winshell
import pyautogui
import time
from ecapture import ecapture as ec
from twilio.rest import Client


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")   

    else:
        speak("Good Evening!")

    assname = ('Smart Speaker.... one point zero.....')    
    speak("I am your Assistant")
    speak(assname)     

def usrname(): 
    speak("What should i call you sir") 
    uname = takeCommand() 
    speak("Welcome Mister") 
    speak(uname) 
    columns = shutil.get_terminal_size().columns 
      
    print("#####################".center(columns)) 
    # print("Welcome Mr.", "\t\t\t\t\t\t\t\t\t",uname)
    print(uname.center(columns))  
    print("#####################".center(columns)) 
      
    speak("How can i Help you, Sir")   

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)    
        print("Say that again please...")  
        return "None"
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('gamersadda.exe@gmail.com', 'Gamers@1234')
    server.sendmail('gamersadda.exe@gmail.com', to, content)
    server.close()

def playmusic(i):
    music_dir = 'F:\\sudhanshu\\songs mp3\\hindi'
    songs = os.listdir(music_dir)
    for (x, item) in enumerate(songs, start=0):
        print(x, '--'+item)   
    os.startfile(os.path.join(music_dir, songs[i]))

def playvideosong(i):
    vid_dir = 'F:\\sudhanshu\\video songs'
    vid_songs = os.listdir(vid_dir)
    for (x, item) in enumerate(vid_songs, start=0):
        print(x, '--'+item) 
    os.startfile(os.path.join(vid_dir, vid_songs[i]))

def NewsFromBBC(): 
      
    # BBC news api 
    main_url = " http://newsapi.org/v2/top-headlines?country=in&sortBy=top&apiKey=2e057248a624436e9ab6aed99c897939"
    # main_url = " https://newsapi.org/v1/articles?source=bbc-news&sortBy=top&apiKey=2e057248a624436e9ab6aed99c897939"
  
    # fetching data in json format 
    open_bbc_page = requests.get(main_url).json() 
  
    # getting all articles in a string article 
    article = open_bbc_page["articles"] 
  
    # empty list which will  
    # contain all trending news 
    results = [] 
      
    for ar in article: 
        results.append(ar["title"]) 
          
    for i in range(len(results)): 
          
        # printing all trending news 
        print(i + 1, results[i]) 
    
    print("do you want me to read them out for you??")
    speak("do you want me to read them out for you??")
    read = takeCommand()
    
    if "yes" or "sure" in read:
        speak("tell me the news number you want me to read")
        no = takeCommand()
        no1 = int(no)
        speak(results[no1])
    else:
        speak("ok Sir, please tell me what can I do next?")
    #to read the news out loud for us 
    # from win32com.client import Dispatch 
    # speak = Dispatch("SAPI.Spvoice") 
    # speak.Speak(results)	

if __name__ == "__main__":
    wishMe()
    usrname() 
    while True:
    # if 1:
        query = takeCommand().lower()

        # Logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            webbrowser.open("stackoverflow.com")   


        elif 'play music' in query or 'play song' in query or 'play a song' in query or 'play the music'\
        in query or 'play the song' in query or 'play some music' in query or 'play some song' in query:
            try:
                speak('please tell song number')
                a = takeCommand()
                i = int(a)
                playmusic(i)
                speak('music is being played')
            except Exception as e:
                print(e)
                speak("sorry, Didn't get the number please try again")

        elif 'play video song' in query or 'play a video song' in query or 'play the video song' in query\
        or 'play some video song' in query:
            try:
                speak('please tell song number')
                a = takeCommand()
                i = int(a)
                playvideosong(i)
                speak('video song is being played')
            except Exception as e:
                print(e)
                speak("sorry Didn't get the number please try again")

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"Sir, the time is {strTime}")

        elif 'open chrome' in query or 'open google chrome' in query:
            chromePath = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
            os.startfile(chromePath)

        elif ' open ms word' in query or "open microsoft word " in query or "open microsoft office word" in query:
            wordPath = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office\\Microsoft Office Word 2007"
            os.startfile(wordPath)
            speak('please wait! while MS word is being opened')

        elif 'send a mail' in query or 'send an email' in query or 'send email' in query: 
            try: 
                speak("What should I say?") 
                content = takeCommand() 
                speak("whome should i send") 
                to = input()     
                sendEmail(to, content) 
                speak("Email has been sent !") 
            except Exception as e: 
                print(e) 
                speak("I am not able to send this email")

        elif 'how are you' in query: 
            speak("I am fine, Thank you") 
            speak("How are you, Sir")

        elif 'am fine' in query or "good" in query: 
            speak("It's good to know that your fine")

        elif 'not fine' in query or "ill" in query or 'not well' in query:  
            speak("I wish I could help you sir")

        elif 'exit' in query: 
            speak("Thanks for giving me your time") 
            exit() 

        elif "who made you" in query or "who created you" in query or "who is your owner" in query or "your owner" in query:  
            speak("I have been created by Mister Sudhanshu Pal Singh.") 

        elif "change name" in query or "change your name" in query: 
            speak("What would you like to call me, Sir ") 
            assname = takeCommand() 
            speak("Thanks for naming me")
            speak(assname)
        
        elif "what's your name" in query or "What is your name" in query or "your name" in query: 
            assname = 'Smart Speaker.... one point zero.....'
            speak("My friends call me") 
            speak(assname) 
            print("My friends call me", assname)

        elif 'joke' in query: 
            speak(pyjokes.get_joke())

        elif 'search' in query or 'play' in query: 
              
            query = query.replace("search", "")  
            query = query.replace("play", "")           
            webbrowser.open(query)

        elif 'change background' in query: 
            ctypes.windll.user32.SystemParametersInfoW(20,  
                                                       0,  
                                                       "D:\\pics\\Freshers Fiesta\\IMG_5716.JPG", 
                                                       0) 
            speak("Background changed succesfully")      



            # http://newsapi.org/v2/top-headlines?country=in&apiKey=

        elif 'news' in query: 
                  
            speak('here are some top news from the times of india') 
            print('''=============== TIMES OF INDIA ============'''+ '\n') 

            NewsFromBBC()
                  
        elif 'lock window' in query or 'lock device' in query or 'lock my pc' in query: 
        	speak("locking the device") 
        	ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query or 'shutdown' in query or 'power off' in query: 
         	speak("Hold On a Sec ! Your system is on its way to shut down") 
         	subprocess.call('shutdown / p /f') 

        elif 'empty recycle bin' in query: 
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True) 
            speak("Recycle Bin Recycled")

        elif "don't listen" in query or "stop listening" in query:
            try:
                speak("for how much Seconds you want to stop jarvis from listening commands")
                i = takeCommand()
                a = int(i)
                speak("ok sir, I will not.... speak....... for")
                speak(i)
                speak("seconds")
                time.sleep(a)
                print(a)
            except Exception as e:
                print(e)
                speak('please provide a valid timing')

        elif "where is" in query: 
            query = query.replace("where is", "") 
            location = query 
            speak("User asked to Locate") 
            speak(location) 
            webbrowser.open("https://www.google.com/maps/place/" + location + "")

        elif "camera" in query or "take a photo" in query: 
            ec.capture(0, "Jarvis Camera ", "img.jpg")

        elif "restart" in query: 
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query: 
            speak("Hibernating") 
            subprocess.call("shutdown / h")

        elif "write a note" in query: 
            speak("What should i write, sir") 
            note = takeCommand() 
            file = open('jarvis.txt', 'w') 
            speak("Sir, Should i include date and time!!! , BYE BYEE!!!") 
            snfm = takeCommand() 
            if 'yes' in snfm or 'sure' in snfm: 
                strTime = datetime.datetime.now().strftime("%m/%d/%Y, %H:%M:%S") 
                file.write(strTime) 
                file.write(" :- ") 
                file.write(note) 
            else: 
                file.write(note)

        elif "weather" in query: 
              
            # Google Open weather website 
            # to get API of Open weather  
            api_key = "70a0ebfa124b45e6fe5eda6b82e6c936" 
            base_url = "https://api.openweathermap.org/data/2.5/weather?"
            speak(" please tell City name ") 
            city_name = takeCommand()
            print("City name :", city_name)  
            complete_url = base_url + "q=" + city_name + "&appid=" + api_key 
            response = requests.get(complete_url)  
            x = response.json()  
              
            if x["cod"] != "404":  
                y = x["main"]  
                current_temperature = y["temp"]  
                current_pressure = y["pressure"]  
                current_humidiy = y["humidity"]  
                z = x["weather"]  
                weather_description = z[0]["description"]  
                print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))  
              
            else:  
                speak(" City Not Found ")


        # elif "close smart speaker" or "close the smart speaker" in query:
        #     try:
        #         speak("BYE BYEE!!! sir, It was nice talking you, remember me whenever you need me")
        #         print("BYE BYEE!!! sir, It was nice talking you, remember me whenever you need me")
        #         break
        #     except Exception as e:
        #         print(e)
        #         speak("please say that again what you want me to do...")

        # elif "send message" or "send a message" in query: 
        #         # You need to create an account on Twilio to use this service 
        #         account_sid = 'ACab83c3965fb59c344997a0793e624307'
        #         auth_token = '14ec7e92b868dc37ab696a7a96c62373'
        #         client = Client(account_sid, auth_token) 
        #         number = "+91"+ str(int(input("enter number: " )))
        #         speak("please enter number")
        #         message = client.messages.create( 
        #                             body = takeCommand(), 
        #                             from_='+12059316078', 
        #                             to = number
        #                         ) 
  
        #         print(message.sid)

        # elif "show note" or "so note" or "show the note" or "so the note" in query: 
        #     speak("Showing Notes")
        #     file = open("jarvis.txt", "r")  
        #     print(file.read()) 
        #     speak(file.read(6))
        #     # path = os.getcwd()
        #     notePath = "D:/sps_documents/python/smart_speak/jarvis.txt"
        #     os.startfile(notePath)
