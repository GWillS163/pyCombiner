# Github: GWillS163
# User: 駿清清 
# Date: 01/11/2022 
# Time: 21:20
# 系统文件操作

import os
import time

from pyCombiner.echos.colorfulPrint import colorPrint


def outputPyFile(restImportLines, restCodeLines, path):
    # make new pyFile to write codes
    try:
        with open(path, 'w', encoding='utf-8') as f:
            for importLine in restImportLines:
                f.write(importLine)
                f.write("\n")
            f.writelines(restCodeLines)
    except PermissionError:
        colorPrint("PermissionError!!!  Change Another Path!", color="red" )
        raise PermissionError


def parseFile(filePath: str) -> list:
    # parse the python file
    # return a list of lines
    with open(filePath, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    return lines


def makeFolder(path):
    if not os.path.exists(path):
        os.makedirs(path)


def getWorkFolderWithFile(entranceFile) -> list:
    """
    Get the folder of entranceFile, and change the workDir to it.
    :param entranceFile:
    :return:
    """
    if not "\\" in entranceFile or not "/" in entranceFile:
        entranceFile = os.path.abspath(entranceFile)
    divideNum = entranceFile.rfind('\\')
    return [entranceFile[:divideNum], entranceFile[divideNum + 1:]]


if __name__ == '__main__':
    outputPyFile(["import os", "import sys"],
                 ["def Print():\n", "\tprint('hello world')\n"],
                 "output\\readMe.py")
