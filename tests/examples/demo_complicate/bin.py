# 4
from service import *


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
