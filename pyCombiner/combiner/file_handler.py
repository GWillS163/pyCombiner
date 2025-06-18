"""
File handling utilities for the Python file combiner
"""

import os
import fnmatch
import chardet
from typing import List, Optional, Set
from pathlib import Path

def find_python_files(source_dir: Path, exclude_patterns: Optional[List[str]] = None) -> List[Path]:
    """Find all Python files in the source directory"""
    if exclude_patterns is None:
        exclude_patterns = []
        
    python_files = []
    exclude_set = set(exclude_patterns)
    
    for root, _, files in os.walk(source_dir):
        for file in files:
            if file.endswith('.py'):
                file_path = Path(root) / file
                relative_path = file_path.relative_to(source_dir)
                
                # Skip if file matches any exclude pattern
                should_exclude = False
                for pattern in exclude_set:
                    if fnmatch.fnmatch(str(relative_path), pattern):
                        should_exclude = True
                        break
                        
                if not should_exclude:
                    python_files.append(relative_path)
                    
    return python_files

def read_file(file_path: Path) -> Optional[str]:
    """Read file contents with proper encoding handling"""
    try:
        # Try UTF-8 first
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except UnicodeDecodeError:
        try:
            # Fall back to latin-1
            with open(file_path, 'r', encoding='latin-1') as f:
                return f.read()
        except Exception as e:
            print(f"Error reading file {file_path}: {e}")
            return None
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return None

def find_python_files_old(directory: str, exclude_patterns: List[str]) -> List[str]:
    """
    递归查找指定目录及其子目录下的所有 Python 文件 (.py)，
    并根据排除模式过滤文件和目录。

    Args:
        directory: 要开始查找的根目录路径。
        exclude_patterns: 要排除的文件或目录的 glob 模式列表。

    Returns:
        符合条件的 Python 文件路径列表 (绝对路径)。
    """
    python_files: List[str] = []
    
    for root, dirnames, filenames in os.walk(directory, followlinks=True):
        # 过滤目录
        dirnames[:] = [
            d for d in dirnames
            if not any(fnmatch.fnmatch(os.path.join(root, d), p) for p in exclude_patterns)
        ]

        for filename in filenames:
            filepath = os.path.join(root, filename)
            if filename.endswith(".py"):
                 # 过滤文件
                if not any(fnmatch.fnmatch(filepath, p) for p in exclude_patterns):
                    python_files.append(os.path.abspath(filepath))

    # 确保文件路径唯一
    return sorted(list(set(python_files)))

def read_file_old(filepath: str) -> str:
    """
    读取指定文件的内容，尝试检测编码并处理可能的解码错误。

    Args:
        filepath: 要读取的文件路径。

    Returns:
        文件内容的字符串。

    Raises:
        FileNotFoundError: 如果文件不存在。
        IOError: 如果读取文件时发生其他 I/O 错误。
        UnicodeDecodeError: 如果文件内容无法用常见编码解码。
    """
    try:
        with open(filepath, 'rb') as f:
            raw_data = f.read()

        # 尝试检测编码
        result = chardet.detect(raw_data)
        encoding = result['encoding'] if result['encoding'] else 'utf-8' # 默认使用 utf-8

        # 使用检测到的编码或默认编码尝试解码
        try:
            return raw_data.decode(encoding)
        except (UnicodeDecodeError, LookupError):
            # 如果检测到的编码失败，尝试一些常见编码
            for enc in ['utf-8', 'latin-1', 'ascii']:
                try:
                    return raw_data.decode(enc)
                except (UnicodeDecodeError, LookupError):
                    continue
            # 所有常见编码都失败，则抛出错误
            raise UnicodeDecodeError(f"无法解码文件 '{filepath}' 的内容。")

    except FileNotFoundError:
        raise FileNotFoundError(f"文件不存在: {filepath}")
    except IOError as e:
        raise IOError(f"读取文件 '{filepath}' 时发生 I/O 错误: {e}")
