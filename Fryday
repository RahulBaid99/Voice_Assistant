import speech_recognition as sr
import os
import sys
import re
import webbrowser
import smtplib
import subprocess
from pyowm import OWM
import wikipedia
from time import strftime
import win32com.client as wincl
speak = wincl.Dispatch("SAPI.SpVoice")

def frydayResponse(audio):
    "speaks audio passed as argument"
    print(audio)
    speak.Speak(audio)

def myCommand():
    "listens for commands"
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print('Say something...')
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source, duration=1)
        audio = r.listen(source)
    try:
        command = r.recognize_google(audio).lower()
        print('You said: ' + command + '\n')
    #loop back to continue to listen for commands if unrecognizable speech is received
    except sr.UnknownValueError:
        print('....')
        command = myCommand();
    return command

def assistant(command):
    "if statements for executing commands"
    #open subreddit Reddit
    if 'open reddit' in command:
        reg_ex = re.search('open reddit (.*)', command)
        url = 'https://www.reddit.com/'
        if reg_ex:
            subreddit = reg_ex.group(1)
            url = url + 'r/' + subreddit
        webbrowser.open(url)
        frydayResponse('The Reddit content has been opened for you Rahul.')

    elif 'stop' in command:
        frydayResponse('Bye bye Rahul. Have a nice day')
        sys.exit()

    #open website
    elif 'open' in command:
        reg_ex = re.search('open (.+)', command)
        if reg_ex:
            domain = reg_ex.group(1)
            print(domain)
            url = 'https://www.' + domain
            webbrowser.open(url)
            frydayResponse('The website you have requested has been opened for you Rahul.')
        else:
            pass

    #greetings
    elif 'hello' in command:
        day_time = int(strftime('%H'))
        if day_time < 12:
            frydayResponse('Hello Rahul. Good morning')
        elif 12 <= day_time < 18:
            frydayResponse('Hello Rahul. Good afternoon')
        else:
            frydayResponse('Hello Rahul. Good evening')

    elif 'help me' in command:
        frydayResponse("""
        You can use these commands and I'll help you out:
        1. Open reddit subreddit : Opens the subreddit in default browser.
        2. Open xyz.com : replace xyz with any website name
        3. Send email/email : Follow up questions such as recipient name, content will be asked in order.
        4. Current weather in {cityname} : Tells you the current condition and temperture
        5. Greetings
        6. news for today : reads top news of today
        7. time : Current system time
        8. top stories from toI news 
        9. tell me about xyz : tells you about xyz
        """)


    #top stories from google news
    elif 'news for today' in command:
        reg_ex = re.search('news (.+)', command)
        if reg_ex:
                domain = 'timesofindia.indiatimes.com'
                url = 'https://www.' + domain
                webbrowser.open(url)
                frydayResponse('Your request has been opened for you Rahul.')
        else:
            pass
        
        

    #current weather
    elif 'current weather' in command:
        reg_ex = re.search('current weather in (.*)', command)
        if reg_ex:
            city = reg_ex.group(1)
            owm = OWM(API_key='*****************')
            obs = owm.weather_at_place(city)
            w = obs.get_weather()
            k = w.get_status()
            x = w.get_temperature(unit='celsius')
            frydayResponse('Current weather in %s is %s. The maximum temperature is %0.2f and the minimum temperature is %0.2f degree celcius' % (city, k, x['temp_max'], x['temp_min']))

    #time
    elif 'time' in command:
        import datetime
        now = datetime.datetime.now()
        frydayResponse('Current time is %d hours %d minutes' % (now.hour, now.minute))

    #send email
    elif 'email' in command:
        frydayResponse('Who is the recipient?')
        recipient = myCommand()
        if 'david' in recipient:
            frydayResponse('What should I say to him?')
            content = myCommand()
            mail = smtplib.SMTP('smtp.gmail.com', 587)
            mail.ehlo()
            mail.starttls()
            mail.login('abc07@gmail.com', '*************')
            mail.sendmail('abc07@gmail.com', 'rahulbaid99@gmail.com', content)
            mail.close()
            frydayResponse('Email has been sent successfuly. You can check your inbox.')
        else:
            frydayResponse('I don\'t know what you mean!')

    
    #ask me anything
    elif 'tell me about' in command:
        reg_ex = re.search('tell me about (.*)', command)
        try:
            if reg_ex:
                topic = reg_ex.group(1)
                ny = wikipedia.page(topic)
                frydayResponse(ny.content[:500].encode('utf-8'))
        except Exception as e:
                print(e)
                frydayResponse(e)

frydayResponse('Hi Rahul, I am Fryday and I am your personal voice assistant, Please give a command or say "help me" and I will tell you what all I can do for you.')

#loop to continue executing multiple commands
while True:
    assistant(myCommand())
    
    ### One Can send mail, launch any app or can check wether by using this commands, although for mail you have to share your id and password and for wether you need to get api keys 
elif 'email' in command:
        frydayResponse('Who is the recipient?')
        recipient = myCommand()if 'rahul' in recipient:
            frydayResponse('What should I say to him?')
            content = myCommand()
            mail = smtplib.SMTP('smtp.gmail.com', 587)
            mail.ehlo()
            mail.starttls()
            mail.login('your_email_address', 'your_password')
            mail.sendmail('sender_email', 'receiver_email', content)
            mail.close()
            frydayResponse('Email has been sent successfuly. You can check your inbox.')else:
            frydayResponse('I don\'t know what you mean!')
            
            elif 'launch' in command:
        reg_ex = re.search('launch (.*)', command)
        if reg_ex:
            appname = reg_ex.group(1)
            appname1 = appname+".app"
            subprocess.Popen(["open", "-n", "/Applications/" + appname1], stdout=subprocess.PIPE)frydayResponse('I have launched the desired application')
            
            
            elif 'current weather' in command:
     reg_ex = re.search('current weather in (.*)', command)
     if reg_ex:
         city = reg_ex.group(1)
         owm = OWM(API_key='ab0d5e80e8dafb2cb81fa9e82431c1fa')
         obs = owm.weather_at_place(city)
         w = obs.get_weather()
         k = w.get_status()
         x = w.get_temperature(unit='celsius')
         frydayResponse('Current weather in %s is %s. The maximum temperature is %0.2f and the minimum temperature is %0.2f degree celcius' % (city, k, x['temp_max'], x['temp_min']))
         ###
