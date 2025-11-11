##################################################################
# SAP Selenium Robot – stabile Ubuntu-Version
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
        """Initialize robot"""
        self.driver = None
        self.wait = None
        print("Robot is ready.")

    def _start_driver_if_not_running(self):
        """Ensure Chrome is started once"""
        if self.driver is not None:
            print("Browser already running.")
            return

        try:
            print("Starting Chrome via webdriver_manager...")
            options = webdriver.ChromeOptions()
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--start-maximized")

            # automatisch passenden ChromeDriver laden
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.wait = WebDriverWait(self.driver, 10)

        except Exception as e:
            print(f"⚠ Chrome konnte nicht gestartet werden: {e}")
            self.driver = None
            raise

    def login(self, username, password):
        """Login in SAP"""
        try:
            self._start_driver_if_not_running()
            if not self.driver:
                raise RuntimeError("WebDriver konnte nicht gestartet werden.")

            login_url = 'https://aimprd.advapp.de/sap/bc/ui2/flp'
            print(f"Öffne Login-Seite: {login_url}")
            self.driver.get(login_url)

            user_field = self.wait.until(
                EC.presence_of_element_located((By.ID, 'USERNAME_FIELD-inner'))
            )
            pass_field = self.driver.find_element(By.ID, 'PASSWORD_FIELD-inner')

            user_field.send_keys(username)
            pass_field.send_keys(password)
            self.driver.find_element(By.ID, 'LOGIN_LINK').click()

            self.wait.until(
                EC.presence_of_element_located((By.ID, 'shell-header-logo'))
            )
            print(f"✅ Login erfolgreich als {username}")
            return {"status": "success", "message": "Login successful"}

        except Exception as e:
            print(f"❌ Login fehlgeschlagen: {e}")
            return {"status": "error", "message": str(e)}

    def quit(self):
        """Schließt den Browser sicher"""
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

robot = SapRobot()

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

        if not json_text.strip():
            continue

        try:
            command_json = json.loads(json_text)
        except json.JSONDecodeError:
            print("⚠ Ungültiges JSON – bitte korrekt eingeben!")
            continue

        action = command_json.get("action")
        params = command_json.get("parameters", {})

        if action == "login":
            result = robot.login(**params)
        elif action == "quit":
            robot.quit()
            result = {"status": "ok", "message": "Browser closed"}
        else:
            result = {"status": "error", "message": f"Unknown action '{action}'"}

        print(f"======== Robot Execution Result ========\n{result}\n==================================")

    except KeyboardInterrupt:
        robot.quit()
        print("\nInterrupted. Exiting...")
        break
    except Exception as e:
        print(f"Unexpected error: {e}")
        robot.quit()
        break