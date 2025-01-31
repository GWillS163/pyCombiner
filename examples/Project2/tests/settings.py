# settings.py
import os

from tests.util.resource import Accounts

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

