# Github: GWillS163
# User: 駿清清 
# Date: 01/11/2022 
# Time: 23:15

def titleBanner(prefix):
    return f"\t------{prefix}------"


def decorate(func):
    def wrapper(prefix, data):
        print(titleBanner(prefix))
        for line in data:
            func(data, line)
        print(titleBanner(prefix))
    return wrapper


@decorate
def showImport(importData, line):
    print(importData[line])


@decorate
def showFrom(fromData, line):
    for importLine in fromData[line].values():
        print(importLine)


@decorate
def showCodes(codes, line):
    print(line)


def showRests(path, imports, froms, codes):
    print("\n\n", path, "-  " * 10)
    showImport("imports", imports)
    showFrom("froms", froms)
    showCodes("codes", codes)


def showMainRests(path, imports, codes):
    print("\n\n", path, "-  " * 10)

    print(titleBanner("restImports"))
    for line in imports:
        print(line)
    print(titleBanner("restImports"))
    # print(titleBanner("code"))
    # for line in codes:
    #     print(line, end="")
    print(f"number of import line: {len(imports)}")
    print(f"number of code lines: {len(codes)}")
    # print(titleBanner("code"))


if __name__ == '__main__':
    importModules = {'update': 'import update as up'}
    fromModules = {'bin': {'*': 'from bin import *'}}
    otherLines = ['', '# 4', 'up.showBanner()', 'main()']
    showRests("Mock Import data", importModules, fromModules, otherLines)
