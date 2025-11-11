# %%
##################################################################
# Cell 1/3: Action Layer (The "Robot")
# Run this cell to "define" your robot in memory
##################################################################
print("Defining Cell 1: SapRobot Action Layer...")

# --- 1. Import all Selenium tools ---
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager  # <--- hinzugefügt
import time

class SapRobot:
    def _init_(self):
        # Kein fester Pfad mehr nötig, wird automatisch über webdriver_manager gelöst
        self.driver = None
        self.wait = None
        print("Robot is ready.")

    def _start_driver_if_not_running(self):
        """Internal function: Check if the browser is running, start it if not"""
        if self.driver is None:
            print("Starting a new browser instance...")

            # WebDriver automatisch verwalten
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service)
            self.driver.maximize_window()
            self.wait = WebDriverWait(self.driver, 10)
        else:
            print("Browser is already running.")

    def login(self, username, password):
        """Action: Log in to the SAP system"""
        try:
            self._start_driver_if_not_running()

            login_url = 'https://aimprd.advapp.de/sap/bc/ui2/flp'
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

            print(f"Successfully logged in as {username}.")
            return {"status": "success", "message": "Login successful"}

        except Exception as e:
            print(f"Login failed: {e}")
            return {"status": "error", "message": str(e)}

    def quit(self):  # <--- hinzugefügt, um den Browser sauber zu schließen
        if self.driver:
            self.driver.quit()
            print("Browser closed successfully.")
        else:
            print("No browser instance to close.")

print("✅ Cell 1/3: SapRobot Action Layer defined.")


# %%
##################################################################
# Cell 3/3: Dispatcher Layer (Modified: Bypassing Ollama)
# Run this cell to "start" your assistant
##################################################################
print("Running Cell 3: Dispatcher Layer (Bypass Ollama Mode)...")

import json
import time

try:
    robot = SapRobot()
except NameError:
    print("="*50)
    print("!! Fatal Error: 'SapRobot' is not defined !!")
    print("="*50)
    raise

print("====================================================")
print("=== SAP Assistant Started (Mode: Direct JSON Command Input) ===")
print("====================================================")
print('Example: {"action": "open_worklist", "parameters": {}}')
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
            print(f"!! Format Error: Your input '{json_text}' is not valid JSON.")
            continue

        action = command_json.get("action")
        params = command_json.get("parameters", {})

        result = None

        if action == "login":
            result = robot.login(**params)

        elif action == "open_worklist":
            result = robot.open_worklist()

        elif action == "quit":
            result = robot.quit()

        else:
            print(f"Error: Received unknown action '{action}'")

        if result:
            print(f"======== Robot Execution Result ========\n{result}\n==================================")

        print("\n--- Waiting for the next command ---")

    except KeyboardInterrupt:
        print("\nInterrupt detected! Closing browser...")
        robot.quit()
        print("Program exited.")
        break
    except Exception as e:
        print(f"\nUnknown error in main loop: {e}")
        robot.quit()
        break