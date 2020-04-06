import os
import pathlib
import requests
from zipfile import ZipFile
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pyautogui
from datetime import datetime, timedelta
import sched
import atexit
from win32com.client import Dispatch
from getpass import getpass

print("Welcome to Class Registration")
print("Run this program 5 minutes before your registration time")
print("Please note this program only works on Windows PCs with either 4K or 1080P screens.")
print("You must have Google Chrome for this to work.")
print("-----------------------------------------------------------------------------------")
print("If you want to just try out this application, enter your time of registration as 3 minutes from "
      "your current time.\nFor example, if it currently is 17:30, enter your registration time as 17:33")
print("-----------------------------------------------------------------------------------")
print("Developed by Rochan :D")
print("-----------------------------------------------------------------------------------")
print()

directory = None
resolution = True
username = None
password = None
passcode = None
crns = None
chrome = None

s = sched.scheduler(time.time, time.sleep)
hour = None
minute = None


def makeDir():
    global directory
    # print("Please enter working directory: ")
    # inp = input()
    inp = os.getcwd()
    directory = pathlib.PureWindowsPath(inp).as_posix()
    if not directory.endswith("/"):
        directory += "/"

    if not os.path.exists(directory):
        os.mkdir(directory)


def getResolution():
    global resolution
    (x, y) = pyautogui.size().width, pyautogui.size().height
    if (x, y) == (3840, 2160):
        resolution = True
    else:
        resolution = False


def getAccountDetails():
    global username, password, passcode, crns, directory
    print("Press '1' for account details set up through a file (advanced)")
    print("Press '2' for manually entering account details (easy)")
    inp = input()
    if inp.__contains__("1"):
        print("Please create a .txt file in the current directory with the following format:")
        print("-----------------------------------------------------------------------------------")
        print("Current Directory: " + os.getcwd())
        print("-----------------------------------------------------------------------------------")
        print("Line 1: GT User ID (Eg:gburdell3)")
        print("Line 2: Password")
        print(
            "Line 3: Duo One Time Passcode (Go to Duo App, Click on Georgia Institute of Technology, and enter the 6 "
            "digit code with no spaces) Eg: 123456")
        print("Line 4: Number of classes you wish to register for: ")
        print("Line 5 to N: Enter the CRNs of the classes you wish to register for separated by a new line")
        print()
        print()
        print("Enter the name of the file:")
        inp = input()
        if not inp.endswith(".txt"):
            inp += ".txt"

        count = 0

        with open(directory + inp) as f:
            for i in range(4):
                line = f.readline()
                count += 1
                if count == 1:
                    username = line.strip()
                if count == 2:
                    password = line.strip()
                if count == 3:
                    passcode = line.strip()
                if count == 4:
                    n = line.strip()
            crns = []
            for i in range(int(n)):
                crns.append(f.readline())

        # crns = remove(crns)
        for i in range(int(n)):
            crns[i] = crns[i].strip()

    elif inp.__contains__("2"):
        username = input("Enter your GT User ID (Eg:gburdell3): ")
        password = getpass("Enter your password: ")
        passcode = input("Enter your Duo One Time Passcode (Go to Duo App, Click on Georgia Institute of Technology, "
                         "and enter the 6 digit code with no spaces) Eg: 123456: ")
        n = int(input("Enter the number of classes you wish to register for: "))
        crns = []
        for i in range(n):
            crns.append(input("Enter the CRN of Class " + str(i + 1) + ": "))


def getChromeVersion(filename):
    parser = Dispatch("Scripting.FileSystemObject")
    try:
        version = parser.GetFileVersion(filename)
    except Exception:
        return None
    return version


def getDriver():
    if os.path.exists(directory + "chromedriver.exe"):
        os.remove(directory + "chromedriver.exe")
    try:
        print("This only works with Google Chrome right now.")
        print("Trying to determine your Chrome Version...")
        paths = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                 r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
        version = list(filter(None, [getChromeVersion(p) for p in paths]))[0]
        print("Your Chrome Version is: " + version)
        v_no = version.split(".")[0]

        if v_no == "79":
            inp = "1"
        elif v_no == "81":
            inp = "3"
        else:
            inp = "2"
    except Exception:
        print(
            "Unable to automatically get your Chrome version. Which version are you using? (Go to Menu -> Help -> "
            "About Google Chrome)")

        print("Enter '1' for Chrome v79.x")
        print("Enter '2' for Chrome v80.x")
        print("Enter '3' for Chrome v81.x")
        inp = input("Selection: ")

    print("Please wait... Downloading selenium driver.")
    if inp.__contains__("1"):
        url = "https://chromedriver.storage.googleapis.com/79.0.3945.36/chromedriver_win32.zip"
        driver_download = requests.get(url)
        open(directory + "driver.zip", 'wb').write(driver_download.content)
    elif inp.__contains__("3"):
        url = "https://chromedriver.storage.googleapis.com/81.0.4044.69/chromedriver_win32.zip"
        driver_download = requests.get(url)
        open(directory + "driver.zip", 'wb').write(driver_download.content)
    else:
        url = "https://chromedriver.storage.googleapis.com/80.0.3987.106/chromedriver_win32.zip"
        driver_download = requests.get(url)
        open(directory + "driver.zip", 'wb').write(driver_download.content)

    with ZipFile(directory + "driver.zip", 'r') as zipfile:
        zipfile.extractall(directory)

    if os.path.exists(directory + "driver.zip"):
        os.remove(directory + "driver.zip")

    print("Done downloading.")


def getRegistrationDetails():
    global hour, minute
    print("Enter your registration time in 24 hour format (HH:MM) [localtime]: ")
    hour, minute = map(int, input().split(":"))
    print()
    print()
    print("Please wait. This application will automatically start 2 minutes before your registration time.")
    print("It will wait at the registration screen and automatically register during your time slot.")
    print("At any moment, please do not exit this application. It will be running in the background.")
    print("Also, please do not try to close the automated Chrome window.")
    print()
    print()


def getReady():
    global chrome

    chrome = webdriver.Chrome(directory + "chromedriver.exe")
    chrome.maximize_window()

    chrome.get("https://buzzport.gatech.edu/cp/home/displaylogin")
    chrome.find_element_by_xpath('//*[@id="panel-log-in"]/div/div/div/a/div[2]/div/i').click()
    chrome.find_element_by_id("username").send_keys(username)
    chrome.find_element_by_id("password").send_keys(password)
    time.sleep(0.5)
    chrome.find_element_by_xpath('//*[@id="login"]/div[5]/input[4]').click()
    time.sleep(0.5)
    if resolution:
        loc = pyautogui.locateCenterOnScreen(directory + "button4k.png")
    else:
        loc = pyautogui.locateCenterOnScreen(directory + "button1008.png")

    pyautogui.moveTo(loc)
    pyautogui.click()
    pyautogui.typewrite(passcode)
    pyautogui.click()
    time.sleep(10)

    chrome.get(
        "https://login.gatech.edu/cas/login?service=https%3A%2F%2Fsso.sis.gatech.edu%3A443%2Fssomanager%2Fc%2FSSB%3Fpkg%3Dtwbkwbis.P_GenMenu?name=bmenu.P_StuMainMnu")
    chrome.get("https://oscar.gatech.edu/pls/bprod/twbkwbis.P_GenMenu?name=bmenu.P_RegMnu")
    chrome.get("https://oscar.gatech.edu/pls/bprod/twbkwbis.P_GenMenu?name=bmenu.P_StuMainMnu")
    chrome.get("https://oscar.gatech.edu/pls/bprod/twbkwbis.P_GenMenu?name=bmenu.P_RegMnu")
    chrome.get("https://oscar.gatech.edu/pls/bprod/bwskfreg.P_AltPin")

    print()
    print("Done loading.")
    print("You may change the term if you wish to now.")


def register():
    global chrome, crns
    chrome.find_element_by_xpath('/html/body/div[3]/form/input').click()

    for i in range(len(crns)):
        crnid = 'crn_id' + str(i + 1)
        chrome.find_element_by_id(crnid).send_keys(crns[i])

    chrome.find_element_by_xpath('/html/body/div[3]/form/input[19]').click()


def main():
    global hour, minute, s, chrome
    makeDir()
    getResolution()
    getAccountDetails()
    getRegistrationDetails()
    getDriver()

    todo_time = (datetime(time.localtime().tm_year, time.localtime().tm_mon, time.localtime().tm_mday,
                          hour, minute, 0) - timedelta(hours=0, minutes=3)).time()

    print()
    print("Waiting until {} to start loading...".format(todo_time))
    s.enterabs(datetime(time.localtime().tm_year, time.localtime().tm_mon, time.localtime().tm_mday, todo_time.hour,
                        todo_time.minute, 0, 0).timestamp(), 1, getReady)

    print("Waiting until {} to register...".format(
        datetime(time.localtime().tm_year, time.localtime().tm_mon, time.localtime().tm_mday, hour, minute, 0).time()))
    s.enterabs(datetime(time.localtime().tm_year, time.localtime().tm_mon, time.localtime().tm_mday, hour,
                        minute, 0, 0).timestamp(), 1, register)

    s.run()


def cleanup():
    global directory
    if os.path.exists(directory + "chromedriver.exe"):
        os.remove(directory + "chromedriver.exe")

    print("Exiting...")


if __name__ == '__main__':
    main()
    atexit.register(cleanup)
