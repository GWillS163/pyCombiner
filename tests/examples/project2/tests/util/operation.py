# operation.py
import time

from tests.util.DriverManager import *
import json


class Color:
    GREEN = '\033[92m'
    RESET = '\033[0m'
    RED = '\033[31m'
    YELLOW = '\033[33m'
    BLUE = '\033[34m'
    MAGENTA = '\033[35m'  # 紫色
    CYAN = '\033[36m'

def makesure_dir_existing(path, path2=None):
    if not os.path.exists(path):
        os.makedirs(path)

    if path2 is not None:
        makesure_dir_existing(f"{path}\\{path2}")

def write_to_result(row, file_path="result.txt", time_stamp=True):
    if time_stamp:
        current_date = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        file_path = file_path.replace(".txt", f"_{current_date}.txt")
    with open(file_path, 'a', encoding='utf-8') as file:
        file.write(row + "\n")

def write_to_result_create(row, file_path="result.txt", time_stamp=True):
    if time_stamp:
        current_date = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        file_path = file_path.replace(".txt", f"_{current_date}.txt")
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(row + "\n")

def longTestDecorator(test_name):
    def decorator(func):
        results_file = f"{test_name}_results.json"
        previous_results = {}

        # 读取之前的测试结果
        if os.path.exists(results_file):
            with open(results_file, "r") as f:
                previous_results = json.load(f)

        def wrapper(buppin_cd):
            failed_tests = []  # 保存失败的测试

            # 筛选未处理的项
            if buppin_cd in previous_results:
                result = previous_results[buppin_cd]
                if result:
                    print(f"跳过已成功的测试: {buppin_cd}")
                    return
                # else:
                #     print(f"跳过已失败的测试: {buppin_cd}")
                #     failed_tests.append(buppin_cd)
                #     return

            # 执行测试
            try:
                func(buppin_cd)  # 调用实际的测试函数
                previous_results[buppin_cd] = True  # 记录成功的测试
                print(f"{Color.GREEN}测试通过: {buppin_cd}{Color.RESET}")
            except Exception as e:
                previous_results[buppin_cd] = False  # 记录失败的测试
                failed_tests.append(buppin_cd)  # 保存失败的测试
                print(f"测试失败: {buppin_cd}，错误: {e}")
                raise e

            # 每次测试后立即保存结果
            with open(results_file, "w") as f:
                json.dump(previous_results, f)

            return failed_tests

        return wrapper

    return decorator



from functools import wraps
from datetime import datetime

def capture_screenshot_on_failure(pause=False, driver=GlobalDriver, filename_template="{function_name}_error_{param}_{times}.png"):
    """
    Decorator to capture a screenshot on failure, dynamically using the 'driver' passed to the test function.
    :param pause:  pause for debugging
    :param driver:
    :param filename_template: Template for the screenshot file name. Use `{param}` to include test parameter.
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            # Ensure the 'driver' is provided as a keyword argument
            # driver = kwargs.get("driver")
            function_name = func.__name__

            if not driver:
                raise ValueError("The 'driver' argument must be provided to the test function.")

            try:
                # Call the decorated function
                func(*args, **kwargs)
            except Exception as e:
                # Generate screenshot file name
                param_value = kwargs.get("times", "")  # Default to 'default' if no param provided
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                # makesure_dir_existing(f"{GlobalDriverManager.get_save_folder()}", f"{function_name}")
                screenshot_path = os.path.join(GlobalDriverManager.get_save_folder(), # function_name,
                                               filename_template.format(function_name=function_name,
                                                                        param=param_value,
                                                                        times=timestamp)
                                               )
                time.sleep(0.5)
                driver.switch_to.window(driver.window_handles[-1])
                driver.save_screenshot(screenshot_path)

                print(f"Error occurred: {e}")
                print(f"Screenshot saved to:\n{screenshot_path}")
                if pause:
                    input("Press Enter to continue...")
                raise e  # Re-raise the exception after handling
        return wrapper
    return decorator

import pytest
from functools import wraps

def xfail_on_exception(exc_type, reason="Expected failure due to exception"):
    """装饰器：捕获异常并标记为预期失败"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except exc_type:
                pytest.xfail(reason)
        return wrapper
    return decorator

class IraiValueError(Exception):
    """Custom exception for inappropriate argument value."""

    def __init__(self, message="依頼無し！ No any Irai appeared here! ", *args):
        super().__init__(message, *args)

def assert_error_style(text):
    return f'\n {"=" * 40}\n {text}\n {"=" * 40}'

if __name__ == "__main__":
    write_to_result("222")