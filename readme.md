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

File Importing Tree：
└── main.py
    ├── module1.py
    │   └── module3.py
    │       └── module4.py
    ├── module2.py
    │   └── module3.py
    │       └── module4.py
    ├── module4.py
    ├── module3.py
    │   └── module4.py
    ├── module4.py
    └── module4.py
```
```Bash
pycombiner main.py -p <project_dir> -s <save_dir>
```


```build
python setup.py sdist build
```
2.4 生成分发档案
下一步是为包生成分发包。这些是上传到包索引的档案，可以通过pip安装。

确保您拥有setuptools并wheel 安装了最新版本：

python3 -m pip install --user --upgrade setuptools wheel
生成这前，可以先运行python setup.py check检查setup.py是否有错误，如果没报错误，则进行下一步输出一般是running check。

1、准备好上面的步骤, 一个包就基本完整了, 剩下的就是打包了，可以使用下面命令打包一个源代码的包:

python setup.py sdist build
这样在当前目录的dist文件夹下, 就会多出一个tar.gz结尾的包了:

2、也可以打包一个wheels格式的包, 使用下面的命令就可以了:

python setup.py bdist_wheel --universal
这样会在dist文件夹下面生成一个whl文件.

3、或者从setup.py位于的同一目录运行此命令：

python3 setup.py sdist bdist_wheel
上面的命令会在dist目录下生成一个tar.gz的源码包和一个.whl的Wheel包。

2.5 发布包到PyPi
1、接下来就是去https://pypi.org/account/register/注册账号，如果有账号的请忽略，记住你的账号和密码，后面上传包会使用。

2、接下来就是上传你的包了，这里使用twine上传。
需要先安装twine（用 twine上传分发包，并且只有 twine> = 1.11.0 才能将元数据正确发送到 Pypi上）

pip install twine
3、安装完之后，运行下面的命令将库上传，上传包，期间会让你输入注册的用户名和密码

twine upload dist/*
输入 PyPI注册的用户名和密码。命令完成后，您应该看到与此类似的输出：

➜  twine upload dist/*
Uploading distributions to https://upload.pypi.org/legacy/
Enter your username: mikezhou_talk
Enter your password:
Uploading package_mikezhou_talk-1.0.0-py3-none-any.whl
100%|██████████████████████████████████████| 7.84k/7.84k [00:03<00:00, 2.29kB/s]
Uploading package_mikezhou_talk-1.0.0.tar.gz
100%|██████████████████████████████████████| 6.64k/6.64k [00:01<00:00, 6.05kB/s]

View at:
https://pypi.org/project/package-mikezhou-talk/1.0.0/
上传完成后，我们的项目就成功地发布到PyPI了。

3.验证发布PYPI成功
上传完成了会显示success,我们直接可以在PyPI上查看，如下:


您可以使用pip来安装包并验证它是否有效。 创建一个新的virtualenv （请参阅安装包以获取详细说明）并从TestPyPI安装包：

python3 -m pip install --index-url https://test.pypi.org/simple/ package-mikezhou-talk
或
pip install package-mikezhou-talk -i https://www.pypi.org/simple/