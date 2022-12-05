# Github: GWillS163
# User: 駿清清 
# Date: 05/12/2022 
# Time: 10:11

# read the inputs from the user in CLI
# and call the main function

import sys
import os.path
from .echos.colorfulPrint import *
from .main import main


def run():
    if len(sys.argv) == 1:
        # print red error message
        colorPrint("No input file", color="red")
        print("Please input the entrance file path.\n"
              "For example: ", end="")
        colorPrint("pyCombiner ./pyCombiner/main.py", color="green")
        sys.exit(0)

    entrancePath = sys.argv[1]
    if not os.path.exists(entrancePath):
        # print red error message
        colorPrint("[Error]", color="red")
        print("No such file or directory: ")
        sys.exit(0)

    outputPath = main(entrancePath)
    print("\nThe output file is: ", end="")
    colorPrint(outputPath, color="green")


if __name__ == '__main__':
    run()
