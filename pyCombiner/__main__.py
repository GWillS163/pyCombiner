# Github: GWillS163
# User: 駿清清 
# Date: 09/11/2022 
# Time: 15:43

# read the inputs from the user in CLI
# and call the main function

import sys
import os.path
from .echos.colorfulPrint import *
from pyCombiner.main import main
from pyCombiner import version, description


def run():
    if len(sys.argv) == 1:
        print(f"Version: {version}\n"
              f"Description: {description}\n"
              " \n")
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
