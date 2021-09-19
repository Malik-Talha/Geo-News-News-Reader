from bs4.element import Script
import requests
import bs4
from requests.api import get
from win32com.client import Dispatch
from win32com.client.makepy import main
import os
    
def clrscr():
    os.system('cls')


def speak(string):
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Speak(string)

def get_url_text(url):
    try:
        return requests.get(url).text
    except Exception as e:
        print("Connection error!")
        print("Plays make sure you have a high speed internet!")
        exit()

        
def getSoup(url_text, parsal):
    return bs4.BeautifulSoup(url_text, parsal)

def getHeadlinesList(soup):
    """ Return list of headlines collected from geo.tv """


    # Add the remaing headlines in tag <h2
    headlines = []
    for h2 in soup.findAll('h2'):
        headlines.append(h2.get_text())

    # Skip the unwanted glitch at the index 0
    headlines = headlines[1:]

    # find the first (featured news) in <a tag at the index 19 
    news = ""
    for a in soup.find_all('a'):
        news += str(a.text)
    list = news.split("\n")
    # Featured news is at the index 19
    headlines.insert(0,list[19])
    return headlines

def speakHeadlines(category, url, headlines, nOfNews):
    clrscr()
    print("News source:\t" + url +"\n")
    if nOfNews == 9:
        speak(f"\t\tTop 10 {category} headlines of today.")
        print(f"\t\tTop 10 {category} headlines of today.")
    elif nOfNews == -1:
        speak(f"\t\tTop {category} headlines of today.")
        print(f"\t\tTop {category} headlines of today.")
    for index, news in enumerate(headlines):
            speak(f"News number {index + 1}!" )
            print(f"{index + 1}.\t{news}")
            speak(news)
            
            if index == nOfNews -1:
                speak("Moving forward to the last news!")
            elif index == nOfNews:
                speak(f"These were the {category} headlines of today")
                clrscr()
                break
            else:
                speak("Moving on to the next news.")


def getPakistanNews():
    url = "https://www.geo.tv/category/pakistan"
    category = "Pakistan"
    soup = getSoup(get_url_text(url), "html.parser")
    headlines = getHeadlinesList(soup)
    print("\nChoose:")
    print("1. Top 10 headlines of today.")
    print("2. All headlines of today.")
    choice = int(input())
    if choice == 1:
        speakHeadlines(category,url, headlines,9)
    elif choice == 2:
        speakHeadlines(category, url, headlines,-1)
    else:
        print("invalid choice")
        

def getWorldNews():
    url = "https://www.geo.tv/category/world"
    category = "World" 
    soup = getSoup(get_url_text(url), "html.parser")
    headlines = getHeadlinesList(soup)
    print("\nChoose:")
    print("1. Top 10 headlines of today.")
    print("2. All headlines of today.")
    choice = int(input())
    if choice == 1:
        speakHeadlines(category, url, headlines,9)
    elif choice == 2:
        speakHeadlines(category, url, headlines,-1)
    else:
        print("invalid choice")


def getSportsNews():
    url = "https://www.geo.tv/category/sports"
    category = "Sports" 
    soup = getSoup(get_url_text(url), "html.parser")
    headlines = getHeadlinesList(soup)
    print("\nChoose:")
    print("1. Top 10 headlines of today.")
    print("2. All headlines of today.")
    choice = int(input())
    if choice == 1:
        speakHeadlines(category, url, headlines,9)
    elif choice == 2:
        speakHeadlines(category, url, headlines,-1)
    else:
        print("invalid choice")

def getLatestNews():
    url = "https://www.geo.tv/latest-news"
    category = "Latest" 
    soup = getSoup(get_url_text(url), "html.parser")
    headlines = getHeadlinesList(soup)
    print("\nchoose:")
    print("1. Top 10 headlines of today.")
    print("2. All headlines of today.")
    choice = int(input())
    if choice == 1:
        speakHeadlines(category, url, headlines,9)
    elif choice == 2:
        speakHeadlines(category, url, headlines,-1)
    else:
        print("invalid choice")

def getEntertainmentNews():
    url = "https://www.geo.tv/category/entertainment"
    category = "Entertainment" 
    soup = getSoup(get_url_text(url), "html.parser")
    headlines = getHeadlinesList(soup)
    print("\nChoose:")
    print("1. Top 10 headlines of today.")
    print("2. All headlines of today.")
    choice = int(input())
    if choice == 1:
        speakHeadlines(category, url, headlines,9)
    elif choice == 2:
        speakHeadlines(category, url, headlines,-1)
    else:
        print("invalid choice")

def getShowbizNews():
    url = "https://www.geo.tv/category/showbiz"
    category = "Showbiz" 
    soup = getSoup(get_url_text(url), "html.parser")
    headlines = getHeadlinesList(soup)
    print("\nChoose:")
    print("1. Top 10 headlines of today.")
    print("2. All headlines of today.")
    choice = int(input())
    if choice == 1:
        speakHeadlines(category, url, headlines,9)
    elif choice == 2:
        speakHeadlines(category, url, headlines,-1)
    else:
        print("invalid choice")

def getTechNews():
    url = "https://www.geo.tv/category/sci-tech"
    category = "Science and Technology" 
    soup = getSoup(get_url_text(url), "html.parser")
    headlines = getHeadlinesList(soup)
    print("\nChoose:")
    print("1. Top 10 headlines of today.")
    print("2. All headlines of today.")
    choice = int(input())
    if choice == 1:
        speakHeadlines(category, url, headlines,9)
    elif choice == 2:
        speakHeadlines(category, url, headlines,-1)
    else:
        print("invalid choice")

def getBusinessNews():
    url = "https://www.geo.tv/category/business"
    category = "Business" 
    soup = getSoup(get_url_text(url), "html.parser")
    headlines = getHeadlinesList(soup)
    print("\nChoose:")
    print("1. Top 10 headlines of today.")
    print("2. All headlines of today.")
    choice = int(input())
    if choice == 1:
        speakHeadlines(category, url, headlines,9)
    elif choice == 2:
        speakHeadlines(category, url, headlines,-1)
    else:
        print("invalid choice")


def getTopNews():
    url = "https://www.geo.tv/"
    category = "Top" 
    soup = getSoup(get_url_text(url), "html.parser")
    headlines = getHeadlinesList(soup)

    # skip a glitch in top category
    headlines = headlines[1:]

    print("\nChoose:")
    print("1. Top 10 headlines of today.")
    print("2. All headlines of today.")
    choice = int(input())
    if choice == 1:
        speakHeadlines(category, url, headlines,9)
    elif choice == 2:
        speakHeadlines(category, url, headlines,-1)
    else:
        print("invalid choice")


if __name__ == "__main__":
    print("Welcome! This Geo news, news reader is developed by Talha Murtaza.")
    speak("Welcome! This Geo news, news reader is developed by Talha Murtaza.")
    while True:
        print("Please Choose one option below:")
        print(f"1. For Latest Headlines:")
        print(f"2. For Pakistan Headlines:")
        print(f"3. For World Headlines:")
        print(f"4. For Top Headlines:")
        print(f"5. For Entertainment Headlines:")
        print(f"6. For Showbiz Headlines:")
        print(f"7. For Sports Headlines:")
        print(f"8. For Technology Headlines:")
        print(f"9. For Business Headlines:")
        print(f"10. To exit:")
        opt = int(input())
        clrscr()
        if opt == 1:
            getLatestNews()
        elif opt == 2:
            getPakistanNews()
        elif opt == 3:
            getWorldNews()
        elif opt == 4:
            getTopNews()
        elif opt == 5:
            getEntertainmentNews()
        elif opt == 6:
            getShowbizNews()
        elif opt == 7:
            getSportsNews()
        elif opt == 8:
            getTechNews()
        elif opt == 9:
            getBusinessNews()
        elif opt == 10:
            print("Thanks for coming!")
            exit()
        else:
            print("Invalid choice! Please choose again.")

    