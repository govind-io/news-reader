from win32com.client import Dispatch
import json
import requests
def say(strng):
    speak=Dispatch("SAPI.SpVoice")
    speak.Speak(strng)
say("Welcome To Live News By Govind")
say("Loading the programm, Checking network")
while True:
    try:
        news=requests.get('https://newsapi.org/v2/top-headlines?sources=bbc-news&apiKey=9474bde79b9b435da10d40f17d329ace')
        news_india=requests.get('http://newsapi.org/v2/top-headlines?country=in&apiKey=9474bde79b9b435da10d40f17d329ace')
        news_india_business=requests.get('http://newsapi.org/v2/top-headlines?country=in&category=business&apiKey=9474bde79b9b435da10d40f17d329ace')
        news_india_entertainment=requests.get('http://newsapi.org/v2/top-headlines?country=in&category=entertainment&apiKey=9474bde79b9b435da10d40f17d329ace')
        news_india_technology=requests.get('http://newsapi.org/v2/top-headlines?country=in&category=technology&apiKey=9474bde79b9b435da10d40f17d329ace')
        break
    except:
        say("You are not connected to the internet please try reconnecting")
        x=input("press enter if you are ready to go \n")
say("status: Connected to network succefully")
parsed_news_india=json.loads(news_india.text)
parsed_news=json.loads(news.text)
parsed_news_india_business=json.loads(news_india_business.text)
parsed_news_india_entertainment=json.loads(news_india_entertainment.text)
parsed_news_india_technology=json.loads(news_india_technology.text)

i=1
say(f"Press 1 for International News Headline) Press 2 for India News Headline, Press 3 for India News Related Technology, Press 4 for India News Entertainment, Press 5 for India News Business ")
while True:
    try:
        choice=int(input())
        while choice>5 or choice<1:
            say("Please enter numbers from 1 to 5")
            print("")
            choice=int(input())
        break
    except ValueError as e:
        say("Please enter integer value only")
if choice==1:
    for item in parsed_news['articles'] :
        if i==1:
            say("Let's Begin With Todays News Headlines")
        say(f"News {i} is:")
        print("Title:"+item['title'])
        say("Title:"+item['title'])
        print('Description:')
        print(item['description'])
        say('description:')
        say(item['description'])
        i=i+1
if choice==2:
    for item in parsed_news_india['articles'] :
        if i==1:
            say("Let's Begin With Todays News Headlines India")
        say(f"News {i} is:")
        print("Title:"+item['title'])
        say("Title:"+item['title'])
        print('Description:')
        print(item['description'])
        say('Description:')
        say(item['description'])
        i=i+1
if choice==5:
    for item in parsed_news_india_business['articles'] :
        if i==1:
            say("Let's Begin With Todays India News related to business.")
        say(f"News {i} is:")
        print("Title:"+item['title'])
        say("Title:"+item['title'])
        print('Description:')
        print(item['description'])
        say('Description:')
        say(item['description'])
        i=i+1
if choice==4:
    for item in parsed_news_india_entertainment['articles'] :
        if i==1:
            say("Let's Begin With Todays India News related to entertainment")
        say(f"News {i} is:")
        print("Title:"+item['title'])
        say("Title:"+item['title'])
        print('Description:')
        print(item['description'])
        say('Description:')
        say(item['description'])
        i=i+1
if choice==3:
    for item in parsed_news_india_technology['articles'] :
        if i==1:
            say("Let's Begin With Todays News Headlines related to technology")
        say(f"News {i} is:")
        print("Title:"+item['title'])
        say("Title:"+item['title'])
        print('Description:')
        print(item['description'])
        say('Description:')
        say(item['description'])
        i=i+1
say("This was all for today. Good bye see you again tomorrow with some more latest and trending news")
