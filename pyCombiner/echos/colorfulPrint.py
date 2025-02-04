# Github: GWillS163
# User: 駿清清 
# Date: 12/11/2022 
# Time: 19:18

def colorPrint(*args, **kwargs):
    """Prints the given arguments in color.

    Args:
        *args: The arguments to print.
        **kwargs: The keyword arguments to pass to the print function.
    """
    colorDict = {
        'black': 30,
        'red': 31,
        'green': 32,
        'yellow': 33,
        'blue': 34,
        'purple': 35,
        'cyan': 36,
        'white': 37,
        'blackBG': 40,
        'redBG': 41,
        'greenBG': 42,
        'yellowBG': 43,
        'blueBG': 44,
        'purpleBG': 45,
        'cyanBG': 46,
    }
    # Get the color from the keyword arguments.
    colorStr = kwargs.pop("color", None)
    colorCode = colorDict.get(colorStr, None)
    # If the color is not None, then print the arguments in color.
    if colorCode:
        print(f"\033[{colorCode}m", end="")
        print(*args, **kwargs)
        print("\033[0m", end="")
    # Otherwise, print the arguments normally.
    else:
        print(*args, **kwargs)


if __name__ == '__main__':

    colorPrint("Hello World!", color="red")
