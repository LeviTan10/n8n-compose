##################################################################
# Cell 1/3: SapRobot Action Layer
##################################################################
print("Defining Cell 1: SapRobot Action Layer...")

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

class SapRobot:
    def _init_(self):
        self.driver = None
        self.wait = None
        print("Robot is ready.")

    def _start_driver_if_not_running(self):
        """Check if Chrome is running, otherwise start it"""
        if self.driver is None:
            print("Starting a new browser instance...")
            service = Service(ChromeDriverManager().install())
            options = webdriver.ChromeOptions()
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--start-maximized")
            self.driver = webdriver.Chrome(service=service, options=options)
            self.wait = WebDriverWait(self.driver, 10)
        else:
            print("Browser is already running.")

    def login(self, username, password):
        """Action: Log in to the SAP system"""
        try:
            # 🔹 Stelle sicher, dass Browser läuft
            self._start_driver_if_not_running()

            login_url = 'https://aimprd.advapp.de/sap/bc/ui2/flp'
            self.driver.get(login_url)

            # 🔹 Warte auf Eingabefelder
            user_field = self.wait.until(
                EC.presence_of_element_located((By.ID, 'USERNAME_FIELD-inner'))
            )
            pass_field = self.driver.find_element(By.ID, 'PASSWORD_FIELD-inner')

            user_field.send_keys(username)
            pass_field.send_keys(password)
            self.driver.find_element(By.ID, 'LOGIN_LINK').click()

            # 🔹 Warte auf Startseite
            self.wait.until(
                EC.presence_of_element_located((By.ID, 'shell-header-logo'))
            )

            print(f"Successfully logged in as {username}.")
            return {"status": "success", "message": "Login successful"}

        except Exception as e:
            print(f"Login failed: {e}")
            return {"status": "error", "message": str(e)}

    def quit(self):
        """Close browser safely"""
        if self.driver:
            self.driver.quit()
            self.driver = None
            print("Browser closed successfully.")
        else:
            print("No browser instance to close.")

print("✅ Cell 1/3: SapRobot Action Layer defined.")


##################################################################
# Cell 3/3: Dispatcher Layer (Bypass Ollama Mode)
##################################################################
print("Running Cell 3: Dispatcher Layer (Bypass Ollama Mode)...")

import json

try:
    robot = SapRobot()
except NameError:
    print("!! Fatal Error: 'SapRobot' is not defined !!")
    raise

print("====================================================")
print("=== SAP Assistant Started (Mode: Direct JSON Command Input) ===")
print("====================================================")
print('Example: {"action": "login", "parameters": {"username": "HSE", "password": "MTproject123"}}')
print('To exit: type q')

while True:
    try:
        json_text = input("Please enter JSON command > ")

        if json_text.lower() == 'q':
            robot.quit()
            print("Program exited.")
            break

        if not json_text:
            continue

        try:
            command_json = json.loads(json_text)
        except json.JSONDecodeError:
            print(f"Invalid JSON: {json_text}")
            continue

        action = command_json.get("action")
        params = command_json.get("parameters", {})

        result = None
        if action == "login":
            result = robot.login(**params)
        elif action == "quit":
            robot.quit()
        else:
            print(f"Unknown action: {action}")

        if result:
            print(f"======== Robot Execution Result ========\n{result}\n==================================")

    except KeyboardInterrupt:
        print("\nInterrupted. Closing browser...")
        robot.quit()
        break
    except Exception as e:
        print(f"Unexpected error: {e}")
        robot.quit()
        break