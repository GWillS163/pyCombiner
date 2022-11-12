# Github: GWillS163
# User: 駿清清 
# Date: 09/11/2022 
# Time: 15:43

# read the inputs from the user in CLI
# and call the main function

import sys
import os.path
from pyCombiner.main import main


print("Welcome to use the CLI of PyCombine! 0.0.1\n")


def run():
    if len(sys.argv) == 1:
        print("Please input the entrance file path.\n"
              "For example: pyCombiner ./pyCombiner/main.py")
        sys.exit(0)

    entrancePath = sys.argv[1]
    if not os.path.exists(entrancePath):
        print("The entrance file path is not correct.")
        sys.exit(0)

    main(entrancePath)
