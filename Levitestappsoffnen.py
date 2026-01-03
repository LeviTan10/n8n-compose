# levitestappsoffnen.py
##################################################################
# Action Layer (SapRobot) - (Adapt Filters Auto-Fix) + Flask Server
##################################################################

print("Defining SapRobot Action Layer (Adapt Filters Auto-Fix)...")

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, WebDriverException

from flask import Flask, request, jsonify
import json
import threading

# ▼▼▼ Global shared driver (keep browser session) ▼▼▼
if 'SHARED_DRIVER' not in globals():
    SHARED_DRIVER = None


class SapRobot:
    def __init__(self):
        global SHARED_DRIVER
        self.driver = None
        self.wait = None
        self.long_wait = None

        # Reconnect to existing browser if possible
        if SHARED_DRIVER is not None:
            try:
                _ = SHARED_DRIVER.window_handles
                print("♻️ Reconnected to existing browser session!")
                self.driver = SHARED_DRIVER
                self._init_waits()
            except:
                print("⚠️ Old browser connection lost. Will start new one.")
                SHARED_DRIVER = None
        else:
            print("🤖 Robot initialized (No active browser found).")

    def _init_waits(self):
        if self.driver:
            self.wait = WebDriverWait(self.driver, 10)
            self.long_wait = WebDriverWait(self.driver, 20)

    def _start_driver_if_not_running(self):
        global SHARED_DRIVER
        if self.driver is not None:
            return

        print("Starting a new browser instance...")
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_experimental_option("detach", True)
        chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self._init_waits()

        SHARED_DRIVER = self.driver
        print("Browser started and saved to global session.")

    def login(self, username, password):
        try:
            self._start_driver_if_not_running()
            login_url = 'https://aimprd.advapp.de/sap/bc/ui2/flp'
            self.driver.get(login_url)

            user_field = self.wait.until(EC.presence_of_element_located((By.ID, 'USERNAME_FIELD-inner')))
            pass_field = self.driver.find_element(By.ID, 'PASSWORD_FIELD-inner')

            user_field.send_keys(username)
            pass_field.send_keys(password)
            self.driver.find_element(By.ID, 'LOGIN_LINK').click()

            self.long_wait.until(EC.presence_of_element_located((By.ID, 'shell-header-logo')))
            print(f"Successfully logged in as {username}.")
            return {"status": "success", "message": "Login successful"}

        except Exception as e:
            return {"status": "error", "message": str(e)}

    # ▼▼▼ Core logic (same flow, only DE/EN text support) ▼▼▼
    def open_app(self, personnel_number, op_plant, order_id, app_name, confirmation_id=None):
        print("Executing workflow: open_app (Tab Strategy)...")
        try:
            self._start_driver_if_not_running()

            # --- Phase 1: Smart navigation ---
            try:
                quick_check = WebDriverWait(self.driver, 3)
                quick_check.until(EC.presence_of_element_located(
                    (By.XPATH, "//input[@placeholder='Search' or @placeholder='Suchen']")
                ))
                print("✅ Already inside app (Search/Suchen visible).")
            except:
                # check personnel screen
                on_personnel_screen = False
                try:
                    WebDriverWait(self.driver, 2).until(EC.presence_of_element_located(
                        (By.XPATH, "//input[contains(@id, 'inputPerNr')]")
                    ))
                    on_personnel_screen = True
                    print("✅ On personnel screen...")
                except:
                    pass

                if not on_personnel_screen:
                    print("Step 0: Go to home...")
                    home_url = 'https://aimprd.advapp.de/sap/bc/ui2/flp#Shell-home'
                    if "Shell-home" not in self.driver.current_url:
                        self.driver.get(home_url)
                    time.sleep(3)

                    print("Step 1: Click tile...")
                    try:
                        self.long_wait.until(EC.element_to_be_clickable((By.ID, "__tile1-title-inner"))).click()
                    except:
                        self.driver.find_element(By.CSS_SELECTOR, "[title='Advanced Manufacturing']").click()
                    time.sleep(5)

            # --- Phase 2: Personnel number ---
            print("Step 2: Check personnel input...")
            try:
                pn_input = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable(
                    (By.XPATH, "//input[contains(@id, 'inputPerNr')]")
                ))
                pn_input.clear()
                pn_input.send_keys(personnel_number)

                # DE/EN submit button
                try:
                    self.driver.find_element(By.XPATH, "//*[text()='Senden']").click()
                except:
                    try:
                        self.driver.find_element(By.XPATH, "//*[text()='Submit']").click()
                    except:
                        pn_input.send_keys(Keys.ENTER)

                print("Personnel number submitted.")
                WebDriverWait(self.driver, 15).until(EC.presence_of_element_located(
                    (By.XPATH, "//input[@placeholder='Search' or @placeholder='Suchen']")
                ))
            except:
                print("No personnel input needed, continuing...")

            # --- Phase 3: Filters (Adapt Filters / Filter anpassen) ---
            print("Step 3: Handle filters...")

            page_source = self.driver.page_source
            if "Op. Plant" not in page_source and "Werk" not in page_source:
                print("❌ 'Op. Plant/Werk' not found -> Adapt Filters / Filter anpassen...")

                self.driver.find_element(
                    By.XPATH,
                    "//*[contains(text(), 'Adapt Filters') or contains(text(), 'Filter anpassen')]"
                ).click()
                time.sleep(2)

                # search box in dialog (DE/EN)
                search_box = self.wait.until(EC.element_to_be_clickable((
                    By.XPATH, "//input[@type='search' or @placeholder='Search' or @placeholder='Suchen']"
                )))
                search_box.clear()
                search_box.send_keys("Op. Plant")
                time.sleep(2)

                # try tick checkbox for Op. Plant
                try:
                    self.driver.find_element(
                        By.XPATH,
                        "//*[text()='Op. Plant']/ancestor::*[contains(@class, 'ListItem') or self::tr]//div[@role='checkbox']"
                    ).click()
                except:
                    # fallback: first checkbox
                    self.driver.find_element(By.XPATH, "//div[@role='checkbox']").click()

                # confirm dialog
                self.driver.find_element(By.XPATH, "//*[text()='OK']").click()
                time.sleep(3)
            else:
                print("✅ Op. Plant/Werk text present, skip Adapt/Filter anpassen.")

            # --- Phase 3.3: Tab strategy input ---
            print("Step 3.3: Keyboard navigation input...")

            search_input = self.wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//input[@placeholder='Search' or @placeholder='Suchen']")
            ))
            search_input.click()
            time.sleep(0.5)

            actions = ActionChains(self.driver)

            print(f"👉 Tab to Order, input: {order_id}")
            actions.send_keys(Keys.TAB).perform()
            time.sleep(0.5)
            self.driver.switch_to.active_element.send_keys(order_id)
            time.sleep(0.5)

            print(f"👉 Tab to Op. Plant/Werk, input: {op_plant}")
            actions.send_keys(Keys.TAB).perform()
            time.sleep(0.5)
            self.driver.switch_to.active_element.send_keys(op_plant)
            self.driver.switch_to.active_element.send_keys(Keys.TAB)
            time.sleep(1)

            # --- Go/Start button (DE/EN) ---
            print("Step 4: Click Go/Start...")
            try:
                self.driver.find_element(By.XPATH, "//button[.='Start']").click()
            except:
                try:
                    self.driver.find_element(By.XPATH, "//button[.='Go']").click()
                except:
                    self.driver.find_element(By.XPATH, "//*[text()='Start' or text()='Go']").click()
            time.sleep(3)

            # 5. Klick Zeile (priorisiere confirmation_id / Rückmeldung)
            if confirmation_id:
                print(f"Step 5: Klick Rückmeldung {confirmation_id}...")
                row = self.long_wait.until(EC.element_to_be_clickable(
                    (By.XPATH, f"//*[text()='{confirmation_id}']/ancestor::tr[1]")
                ))
            else:
                print(f"Step 5: Klick Auftrag {order_id}...")
                row = self.long_wait.until(EC.element_to_be_clickable(
                    (By.XPATH, f"//*[text()='{order_id}']/ancestor::tr[1]")
                ))

            row.click()
            time.sleep(2)


            # --- Open final app ---
            print(f"Step 6: Open App '{app_name}'...")
            try:
                self.long_wait.until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, f"[title='{app_name}']")
                )).click()
            except:
                print("Title match failed -> text match...")
                short_name = app_name.split(' ')[0]
                self.driver.find_element(By.XPATH, f"//*[contains(text(), '{short_name}')]").click()

            print(f"Clicked '{app_name}' Tile!")
            return {"status": "success", "message": f"Workflow complete: {app_name}"}

        except Exception as e:
            print(f"!! Workflow failed: {e}")
            import traceback
            traceback.print_exc()
            return {"status": "error", "message": str(e)}

    def quit(self):
        if self.driver:
            self.driver.quit()
            self.driver = None


print("✅ SapRobot Ready.")
print("Running Robot Server...")

app = Flask(__name__)
robot = SapRobot()
print("Robot is ready and waiting for commands.")


def dispatch_action(data_wrapper: dict):
    print(f"Received data wrapper for dispatch: {data_wrapper}")
    try:
        # ✅ FIX: accept BOTH formats
        # A) direct: {"action":"login","params":{...}}
        # B) wrapped: {"output":"...{...}..."}
        if isinstance(data_wrapper, dict) and "action" in data_wrapper:
            command_json = data_wrapper
        else:
            output_string = data_wrapper.get("output")
            if not output_string:
                print("!! Error: 'output' key is missing or empty.")
                return

            start_index = output_string.find('{')
            end_index = output_string.rfind('}')
            if start_index == -1 or end_index == -1:
                print(f"!! Error: Could not find JSON {{...}} in string: {output_string}")
                return

            json_string_cleaned = output_string[start_index:end_index + 1]
            print(f"Cleaned JSON string: {json_string_cleaned}")
            command_json = json.loads(json_string_cleaned)

        action = command_json.get("action")
        params = command_json.get("params", {}) or {}

        print(f"Dispatching action: {action} with params: {params}")

        result = None
        if action == "login":
            result = robot.login(**params)

        elif action == "open_app":
            safe_params = {
                "personnel_number": params.get("personnel_number"),
                "op_plant": params.get("op_plant"),
                "order_id": params.get("order_id"),
                "app_name": params.get("app_name"),
                "confirmation_id": params.get("confirmation_id")
            }
            safe_params = {k: v for k, v in safe_params.items() if v is not None}
            result = robot.open_app(**safe_params)

        elif action == "quit":
            result = robot.quit()

        else:
            result = {"status": "error", "message": f"Unknown action: {action}"}

        print(f"Action '{action}' finished with result: {result}")

    except Exception as e:
        print(f"!! Error during action dispatch: {e}")


@app.route('/execute', methods=['POST'])
def execute_command():
    print("\n--- Received command from n8n! ---")
    try:
        raw_data = request.data
        if not raw_data:
            print("!! Error: Received empty request body.")
            return jsonify({"status": "error", "message": "Received empty request body"}), 400

        json_string = raw_data.decode('utf-8')
        print(f"Received raw text: {json_string}")
        command_data = json.loads(json_string)

        thread = threading.Thread(target=dispatch_action, args=(command_data,))
        thread.start()

        return jsonify({"status": "received", "message": "Command received and is being processed."})
    except Exception as e:
        print(f"!! Error in execute_command: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


if __name__ == "__main__":
    print("====================================================")
    print("=== Python Robot Server is starting... ===")
    print("=== Listening on http://localhost:5000/execute ===")
    print("====================================================")
    app.run(host='localhost', port=5000, debug=False, use_reloader=False)
