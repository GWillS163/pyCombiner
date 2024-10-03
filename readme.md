# The Goal you want
- If you want to create a single Python script(.py) from Python code that you have written, you should try **pyCombiner**.
- If you hope to create a single Python script(.py) that from Python packages , Perhaps you should try **stickytape**.
- If you want to create a single file(.pyz) that can be executed by a Python interpreter, use **zipapp**.
- If you need to create a standalone executable(.exe) from your Python script, I recommend using an alternative such as **PyInstaller**.

# The Function of PyCombiner

it is a commandline tool to combine python files into one file.
The generated resultant and usage is like pyinstaller, but is much simpler, and resultant is not ".exe", is ".py".

path on this repository:
examples/demo_complicate
<img src="https://github.com/GWillS163/pyCombiner/raw/master/res/introImg.png">

# Install 
```Bash
pip install pycombiner
```

# Example
```Bash
# Project Folder"
├ main.py
└─src_lib
        __init__.py
        os_lib.py
```
```Bash
pycombiner main.py
```
