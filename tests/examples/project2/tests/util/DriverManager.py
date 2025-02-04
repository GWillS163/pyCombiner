# DriverManager.py
import os
from datetime import datetime
from selenium import webdriver
from tests.settings import settings


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