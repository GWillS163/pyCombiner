# 2
import time,os


def isLoginValid(userInfo, adminPass):
    return userInfo and adminPass


def getInfo():
    return "Info"


def colorPrint(color, string, end="\n"):
    # print string with 33m
    colorSigns = {
        "red": "31m",
        "green": "32m",
        "yellow": "33m",
        "blue": "34m",
        "purple": "35m",
        "cyan": "36m",
        "white": "37m",
        "black": "30m"
    }
    print("\033[", colorSigns[color], string, "\033[0m", sep="", end=end)


def showNamePrompt():
    colorPrint("purple", "Enter your name: ")


def showPasswordPrompt():
    colorPrint("purple", "Enter your password: ")


def inputSign():
    colorPrint("green", ">>>", end=" ")
    return input()
