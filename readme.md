![5CE63D7A](https://github.com/user-attachments/assets/f025e89d-a996-4715-ba74-dfebed360b4d)
## 🌐 Project Overview | 简介 | 概要

**English**

> `pyCombiner` is a lightweight CLI tool that merges multiple Python modules into a single `.py` file. It's ideal for simplifying distribution and debugging, without the need for converting to `.exe`.

**中文**

> `pyCombiner` 是一个轻量级命令行工具，用于将多个 Python 模块合并为单一文件，便于发布和调试，无需转换成 `.exe`。

**日本語**

> `pyCombiner`は複数の Python ファイルを 1 つの `.py`ファイルにまとめるコマンドラインツールです。

---

## 🚀 The Goal You Want | 功能比较 | 比較

| Use Case                   | Tool           | Output Type       |
| -------------------------- | -------------- | ----------------- |
| Merge self-written modules | **pyCombiner** | Single `.py` file |
| Merge installed packages   | `stickytape`   | Single `.py` file |
| Executable zip             | `zipapp`       | `.pyz` file       |
| Create `.exe`              | `PyInstaller`  | `.exe` binary     |

---

## ⚡ Function of pyCombiner | 功能介绍 | 機能

`pyCombiner` is a command-line utility that recursively merges Python files into a single script.

* Similar to PyInstaller but simpler
* Result is `.py` instead of `.exe`

**Demo Project**:
`examples/demo_complicate`

![demo image](https://github.com/GWillS163/pyCombiner/raw/master/res/introImg.png)

---

## ⚙️ Install | 安装 | インストール

```bash
pip install pycombiner
```

---

## 📙 Quick Start | 快速上手 | クイックスタート

You can try `pyCombiner` using the built-in demo project:

**Step 1: Clone the demo project**

```bash
git clone https://github.com/GWillS163/pyCombiner.git
cd pyCombiner/tests/examples/deep_demo
```

**Step 2: Run pyCombiner on the example**

```bash
pycombiner main.py
```

**Step 3: Check the output**
```bash
PS C:\xxxxx\pyCombiner\tests\examples\deep_demo> pycombiner main.py
開始ファイルパス: main.py
プロジェクトディレクトリ: C:\xxxxx\pyCombiner\tests\examples\deep_demo
保存ディレクトリ: C:\xxxxx\pyCombiner\tests\examples\deep_demo
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

The output file is: C:\xxxxx\pyCombiner\tests\examples\deep_demo\main20250602_172431_combined.py
```
🤩 Done!

**Output:**
```bash
dir  # or 'ls' on Unix-based systems
```
```bash
    ディレクトリ: C:\xxxxx\pyCombiner\tests\examples\deep_demo

Mode                 LastWriteTime         Length Name
----                 -------------         ------ ----
d-----        2025/06/02     17:23                subdir
-a----        2025/06/02     17:23            270 main.py
-a----        2025/06/02     17:24            308 main20250602_172431_combined.py
-a----        2025/06/02     17:23             62 module1.py
-a----        2025/06/02     17:23              0 __init__.py

```

---

## ✨ Credits & License

* Created by [GWillS163](https://github.com/GWillS163)
* Open source under [MIT License](LICENSE)
