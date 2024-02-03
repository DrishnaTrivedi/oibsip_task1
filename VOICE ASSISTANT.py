import win32com.client
import speech_recognition as sr
import time
from datetime import date
import datetime
import wikipedia
import windowsapps
from googlesearch import search
import webbrowser

t=time.strftime("%H : %M ")
speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Voice = speaker.GetVoices().Item(1) #changing voice
speaker.rate=2

def takecommand():

    speaker.Speak("Hello, I'm your voice assistant! From telling you the time, date, and day, to searching the web or opening apps – just say the word, and let the magic unfold!")
    print("listening...")

    r= sr.Recognizer()

    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source,phrase_time_limit=5)

        try:
            query = r.recognize_google(audio,language ="en-in")
            print(query)

            if("Hi" in query.split() or "Hello" in query.split() or "hi" in query.split() or "hello" in query.split() ):
                speaker.Speak("hii there! Great to have you!!")

            elif("time" in query.split()):
                speaker.Speak("The current time is ")
                speaker.Speak(t)

            elif ("date" in query.split() and "day" in query.split()):
                date1 = date.today()
                speaker.Speak(f"Todays date is {date1}")

                speaker.Speak("and day is")
                speaker.Speak(time.strftime("%A"))

            elif("date" in query.split()):
                date1= date.today()
                speaker.Speak(f"Todays date is {date1}")


            elif("day" in query.split()):
                day_of_week = time.strftime("%A")
                speaker.Speak(f"Today is {day_of_week}")


            elif("search" in query.split()):
                query = query.replace("search for", "")
                try:
                    for j in search(query, num=10, stop=10, pause=2):
                        print(j)
                    for j in search(query, num=1, stop=1):
                        webbrowser.open_new_tab(j)
                    try:
                        # query = query.replace("search for", "")
                        search_results = wikipedia.search(query)
                        summary = wikipedia.summary(search_results[0], sentences=1)
                        print(summary)
                        speaker.Speak(summary)

                        # for j in search(query, num=10, stop=10, pause=2):
                        #        print(j)
                        # for j in search(query, num=1, stop=1):
                        #     webbrowser.open_new_tab(j)

                    except:
                        print("Couldn't fetch results")

                except:
                    print("Couldn't fetch results.check your internet connectivity")
                # try:
                #     query = query.replace("search for", "")
                #     search_results = wikipedia.search(query)
                #     summary = wikipedia.summary(search_results[0], sentences=2)
                #     print(summary)
                #     speaker.Speak(summary)
                # except:
                #     print("Couldn't fetch results")


            elif ("open" in query.split()):
                # query = query.replace("open", "")
                try:
                    windowsapps.open_app(query[5:])
                    speaker.Speak("There you go!!")
                except:
                    speaker.Speak("There is no such application on the device")

        except sr.UnknownValueError:
            print("Sorry, I could not understand the audio.")
            speaker.Speak("Sorry, I could not understand the audio.")

        except sr.RequestError as e:
            print(f"Could not request results from Google Speech Recognition service; {e}")
            speaker.Speak("Could not request results from Google Speech Recognition service. Check your internet connectivity")


if __name__ == '__main__':
    takecommand()