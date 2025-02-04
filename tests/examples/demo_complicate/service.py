# 3
from utils import *


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
