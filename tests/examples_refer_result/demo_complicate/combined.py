# 5
# 1

welcomeBanner = "Welcome use Version 1.0.0"
# 4
# 3
# 2
import time,os
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


def getName():
    """
    Get user info
    :return:
    """
    showNamePrompt()
    return inputSign()


def getPasswd():
    """
    Get admin password
    :return:
    """
    showPasswordPrompt()
    return inputSign()


def showBanner(bannerStr):
    """
    Show banner
    :param bannerStr:
    :return:
    """
    colorPrint("cyan", bannerStr)


def main(welcomeBanner):
    """
    CN: 如果用户输入的两个值不为空，则显示登录成功
    EN: if adminName and adminPwd that user input are not empty both, then login success
    :param welcomeBanner:
    :return:
    """
    showBanner(welcomeBanner)
    adminName = getName()
    adminPass = getPasswd()

    if isLoginValid(adminName, adminPass):
        colorPrint("green", "Login Success")
    else:
        colorPrint("red", "Login Failed")

# CN: 只是一个demo
# EN: Just a demo_complicate
main(welcomeBanner)
