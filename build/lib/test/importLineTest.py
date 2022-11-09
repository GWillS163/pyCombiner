# Github: GWillS163
# User: 駿清清 
# Date: 01/11/2022 
# Time: 21:51
from pprint import pprint


def singleLineParseTest():
    importLines = [
        "import os as o",
        "import re, utils"
    ]

    fromLines = [
        "from os import path, sys",
        "from os import path as p"
    ]

    for importLine, fromLine in zip(importLines, fromLines):
        print("Import:", end="\t")
        print(getImportLineDict(importLine))
        print("From:", end="\t")
        pprint(getFromLineDict(fromLine))


def singlePyTest(path):
    from pyCombiner import parseFile
    from pyCombiner import parseSinglePy
    from pyCombiner import showRests

    lines = parseFile(path)
    importModules, fromModules, otherLines = parseSinglePy(lines)
    showRests(path, importModules, fromModules, otherLines)


if __name__ == '__main__':
    # singleLineParseTest()
    singlePyTest("../examples/demo_simple/runMeSim.py")
    # singlePyTest("../demo_simple/utils.py")
    singlePyTest("../examples/demo_simple/update.py")
