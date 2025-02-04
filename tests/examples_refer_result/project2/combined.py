import os
from datetime import datetime
import psycopg2
from psycopg2 import sql
# ========================  BEGIN inlined module: tests.util.operation
# operation.py
import time

# ========================  BEGIN inlined module: tests.util.DriverManager
# DriverManager.py
import os
from datetime import datetime
from selenium import webdriver
# ========================  BEGIN inlined module: tests.settings
# settings.py
import os

# ========================  BEGIN inlined module: tests.util.resource
class Accounts:
    credentials = {
        "本庁": ["10000", "1000010000"],
        "艦艇装備研究所": ["16000", "1600016000"],
        "次世代装備研究所": ["19000", "1900019000"]
    }

    @classmethod
    def get_credentials(cls, name) -> list:
        pair = cls.credentials.get(name)
        if pair is None:
            raise ValueError(f"[ {name} ] not existing")
        return pair # usr, pwd

class Kanri_irai:
    # 【依頼】
    供用_0 = 0
    返納_1 = 1

    # 【受入依頼】
    寄付_2 = 2
    借受_3 = 3
    編入_4 = 4
    転用_5 = 5
    雑件_6 = 6

    # 【生産報告】
    生産_7 = 7
    副生_8 = 8
    発生材_9 = 9
    物品減耗_10 = 10

    # 【不用決定依頼】
    転用_不用_11 = 11
    売払_12 = 12
    解体_13 = 13
    廃棄_14 = 14

    # 【払出依頼】
    交換_15 = 15
    編入_払出_16 = 16
    譲与_17 = 17
    貸付_18 = 18
    亡失_19 = 19
    雑件_払出_20 = 20

    # 【寄託依頼】
    寄託_21 = 21
    物品保管_22 = 22

    # 【管理換】
    管理換_払_23 = 23
    管理換_受_24 = 24
    分類換_払_25 = 25
    分類換_受_26 = 26
    供用換_払_27 = 27
    管理換協議依頼_払_28 = 28
    管理換協議依頼_受_29 = 29
    管理換協議通知_払_30 = 30
    管理換協議通知_受_31 = 31

    # 【依頼差戻し】
    依頼差戻し_32 = 32


class Kyouyou_irai:
    # 【管理官取得】
    新規取得_0 = 0
    返品_材料使用_1 = 1
    交換_2 = 2
    管理換_受_3 = 3
    分類換_受_4 = 4
    管理官取得_雑件_5 = 5

    # 【受入】
    # 【供用受入】
    供用_6 = 6
    供用換_7 = 7
    # 【生産受入】
    生産_8 = 8
    副生_9 = 9
    発生材_10 = 10
    # 【受入】
    寄付_11 = 11
    借受_12 = 12
    編入_13 = 13
    転用_14 = 14
    受入_雑件_15 = 15
# ========================  END inlined module: tests.util.resource

virtual_machine = "http://10.216.192.185/tomcat/"

public_machine = "http://192.168.11.109/tomcat"
public_old_machine = "http://192.168.11.170/tomcat"
local_tomcat = "http://localhost:8080/tomcat"
local_alot = "http://localhost:8080/alot"


project_path = r"C:\Users\PFS\PycharmProjects\AutoUITest"

test_bango = "TEST-2025"

class settings:
    base_url = virtual_machine
    usr, pwd = Accounts.get_credentials("本庁")
    full_test_name = "Tomcat.war-1218Test"

    # Take ScreenShot
    save_screenshot = os.getenv("SCREENSHOT", "false").lower() == "true"
    save_to_screen_wikipedia = True
    wikipedia_path = f"{project_path}\\tests\\screen_wikipedia"
    save_to_test_record = True
    test_record_path = f"{project_path}\\tests\\output"

    # Browser
    headless_mode = os.getenv("HEADLESS_MODE", "false").lower() == "true"

print("Using: ", settings.usr, settings.pwd)

# ========================  END inlined module: tests.settings


class DriverManager:
    _instance = None
    screen_save_folder = None

    @classmethod
    def create_save_folder(cls):
        now = datetime.now()
        cls.screen_save_folder = f"{settings.test_record_path}\\{now.strftime('%Y-%m-%d_%H-%M-%S')}"
        if not os.path.exists(cls.screen_save_folder):
            os.makedirs(cls.screen_save_folder)
        return cls.screen_save_folder

    @classmethod
    def get_save_folder(cls):
        return cls.screen_save_folder

    @classmethod
    def get_driver(cls):
        if cls._instance is None:
            cls.screen_save_folder = cls.create_save_folder()
            options = webdriver.ChromeOptions()

            if settings.headless_mode:
                options.add_argument("--headless")
                options.add_argument("--disable-gpu")
            # options.add_argument('--no-sandbox')  # 可选：在某些环境下提高稳定性
            # options.add_argument('--disable-dev-shm-usage')  # 可选：减少内存使用
            cls._instance = webdriver.Chrome(options=options)  # 使用无头选项
        return cls._instance

    @classmethod
    def quit_driver(cls):
        if cls._instance is not None:
            cls._instance.quit()
            cls._instance = None

GlobalDriverManager = DriverManager()
GlobalDriver = GlobalDriverManager.get_driver()
# ========================  END inlined module: tests.util.DriverManager
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

# if __name__ == "__main__":
#     write_to_result("222")
# ========================  END inlined module: tests.util.operation


class DatabaseSearcher:
    def __init__(self, host, dbname, user, password, key_name, values):
        self.host = host
        self.dbname = dbname
        self.user = user
        self.password = password
        self.key_name = key_name
        self.values = values
        self.conn = None
        self.cur = None

        # 设置输出文件路径
        current_date = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        self.output_dir = f"Output/DB [{self.key_name}]_{current_date}-[" + self.values.replace('%', '').replace('\t','')[:10] + "]/"
        os.makedirs(self.output_dir, exist_ok=True)  # 如果文件夹不存在，创建文件夹

    def connect(self):
        """连接数据库"""
        self.conn = psycopg2.connect(
            host=self.host, dbname=self.dbname, user=self.user, password=self.password
        )
        self.cur = self.conn.cursor()

    def disconnect(self):
        """关闭数据库连接"""
        if self.cur:
            self.cur.close()
        if self.conn:
            self.conn.close()

    def output_file_print(self, results, table_name):
        """打印和写入查询结果"""
        file_path = os.path.join(self.output_dir, f"{table_name}_({len(results)} items).txt")

        print(f"\t{Color.GREEN}Found matching rows in {table_name}{Color.RESET}:")

        for i, row in enumerate(results):
            if not i > 10:
                print(row)

            write_to_result(str(row), file_path)

        print(f"------Total ({len(results)}) items\n")

    def search(self):
        """搜索数据库中的表和记录"""
        try:
            # 获取所有表的名字
            self.cur.execute("""
                SELECT table_name 
                FROM information_schema.tables 
                WHERE table_schema = 'admin'
            """)
            tables = self.cur.fetchall()

            for table in tables:
                table_name = table[0]
                if "view" in str(table_name):
                    continue

                print(f"Checking table: {table_name}", end=" ")

                # 检查当前表是否有指定的列
                self.cur.execute("""
                    SELECT column_name
                    FROM information_schema.columns
                    WHERE table_name = %s AND column_name = %s
                """, (table_name, self.key_name))

                columns = self.cur.fetchall()

                if columns:
                    print(
                        f"\nFound '{self.key_name}' in table: {Color.MAGENTA}{table_name: ^20}{Color.RESET}. Querying for {self.values}...",
                        end=" "
                    )
                    if "%" in self.values:
                        query = sql.SQL("SELECT * FROM {} WHERE {} LIKE %s").format(
                            sql.Identifier(table_name),  # 动态插入表名
                            sql.Identifier(self.key_name)  # 动态插入列名
                        )
                    else:
                        query = sql.SQL("SELECT * FROM {} WHERE {} = %s").format(
                            sql.Identifier(table_name),  # 动态插入表名
                            sql.Identifier(self.key_name)  # 动态插入列名
                        )
                    self.cur.execute(query, (self.values,))
                    results = self.cur.fetchall()

                    if results:
                        self.output_file_print(results, table_name)
                    else:
                        print(f"\t{Color.RED}[No matching rows found]{Color.RESET}\n")
                else:
                    print(f"\t\t{Color.RED}[No '{self.key_name}' column]{Color.RESET}\n")

        finally:
            self.disconnect()


# 示例调用
# if __name__ == "__main__":
#     host = 'localhost'
#     dbname = 'bks'
#     user = 'admin'
#     password = 'admin'
#     key_name = 'kanribo_cd_harai'
#     values = '1952848' # if "%" in

#     searcher = DatabaseSearcher(host, dbname, user, password, key_name, values)
#     searcher.connect()
#     searcher.search()
