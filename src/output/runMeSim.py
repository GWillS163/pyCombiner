import os
# 1
def showBanner():
    print('=' * 40)
    print('Welcome to the system')
    print('=' * 40)
# 2

def showLogin():
    print('=' * 40)
    print("Login success")
    print('=' * 40)
    print(os.path.abspath(__file__))


def showExit():
    print('=' * 40)
    # show menu
    print('Exited')
    print('=' * 40)

# 3
def main():
    print("Hello World!")
    showLogin()
    showExit()

# 4
up.showBanner()
main()
