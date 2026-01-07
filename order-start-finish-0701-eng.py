# %%
##################################################################
# Cell 1/3: Action Layer (With Start/End Order Logic)
##################################################################

print("Defining SapRobot Action Layer (With Start/End Order Logic)...")

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

# ▼▼▼ Global Variables ▼▼▼
if 'SHARED_DRIVER' not in globals():
    SHARED_DRIVER = None


class SapRobot:
    def __init__(self):
        global SHARED_DRIVER
        self.driver = None
        self.wait = None
        self.long_wait = None

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

    # ▼▼▼ Core Navigation Logic (Unchanged) ▼▼▼
    def open_app(self, personnel_number, op_plant, order_id, app_name, confirmation_id=None):
        print(f"Executing workflow: open_app for '{app_name}'...")
        try:
            self._start_driver_if_not_running()

            # --- Phase 1: Smart Navigation ---
            try:
                quick_check = WebDriverWait(self.driver, 3)
                quick_check.until(EC.presence_of_element_located(
                    (By.XPATH, "//input[@placeholder='Search' or @placeholder='Suchen']")
                ))
                print("✅ Already inside app (Search/Suchen visible).")
            except:
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

            # --- Phase 2: Personnel Number ---
            print("Step 2: Check personnel input...")
            try:
                pn_input = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable(
                    (By.XPATH, "//input[contains(@id, 'inputPerNr')]")
                ))
                pn_input.clear()
                pn_input.send_keys(personnel_number)

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

            # --- Phase 3: Filters (Adapt Filters) ---
            print("Step 3: Handle filters...")
            page_source = self.driver.page_source
            if "Op. Plant" not in page_source and "Werk" not in page_source:
                print("❌ 'Op. Plant/Werk' not found -> Adapt Filters...")
                self.driver.find_element(
                    By.XPATH,
                    "//*[contains(text(), 'Adapt Filters') or contains(text(), 'Filter anpassen')]"
                ).click()
                time.sleep(2)

                search_box = self.wait.until(EC.element_to_be_clickable((
                    By.XPATH, "//input[@type='search' or @placeholder='Search' or @placeholder='Suchen']"
                )))
                search_box.clear()
                search_box.send_keys("Op. Plant")
                time.sleep(2)

                try:
                    self.driver.find_element(
                        By.XPATH,
                        "//*[text()='Op. Plant']/ancestor::*[contains(@class, 'ListItem') or self::tr]//div[@role='checkbox']"
                    ).click()
                except:
                    self.driver.find_element(By.XPATH, "//div[@role='checkbox']").click()

                self.driver.find_element(By.XPATH, "//*[text()='OK']").click()
                time.sleep(3)
            else:
                print("✅ Op. Plant/Werk text present, skip Adapt.")

            # --- Phase 3.3: Keyboard Input Strategy ---
            print("Step 3.3: Keyboard navigation input...")
            search_input = self.wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//input[@placeholder='Search' or @placeholder='Suchen']")
            ))
            search_input.click()
            time.sleep(0.2)
            search_input.send_keys(Keys.CONTROL, "a")
            search_input.send_keys(Keys.DELETE)
            time.sleep(0.2)

            actions = ActionChains(self.driver)
            print(f"👉 Tab to Order, input: {order_id}")
            actions.send_keys(Keys.TAB).perform()
            time.sleep(0.4)
            self.driver.switch_to.active_element.send_keys(order_id)

            print("Step 4: Click Go/Start...")
            try:
                self.driver.find_element(By.XPATH, "//button[.='Start']").click()
            except:
                try:
                    self.driver.find_element(By.XPATH, "//button[.='Go']").click()
                except:
                    self.driver.find_element(By.XPATH, "//*[text()='Start' or text()='Go']").click()
            time.sleep(3)

            # 5. Click Order Row
            if confirmation_id:
                print(f"Step 5: Click Confirmation {confirmation_id}...")
                row = self.long_wait.until(EC.element_to_be_clickable(
                    (By.XPATH, f"//*[text()='{confirmation_id}']/ancestor::tr[1]")
                ))
            else:
                print(f"Step 5: Click Order {order_id}...")
                row = self.long_wait.until(EC.element_to_be_clickable(
                    (By.XPATH, f"//*[text()='{order_id}']/ancestor::tr[1]")
                ))
            row.click()
            time.sleep(2)

            # 6. Open Final App (If this is the target)
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
            return {"status": "success", "message": f"Opened: {app_name}"}

        except Exception as e:
            print(f"!! Workflow failed: {e}")
            import traceback
            traceback.print_exc()
            return {"status": "error", "message": str(e)}

    # ▼▼▼ 🟢 Fixed Version: Support Specific Confirmation ID (Operation) ▼▼▼
    def process_order_step(self, personnel_number, op_plant, order_id, phase, action_type, confirmation_id=None):
        """
        action_type: 'start' or 'end'
        confirmation_id: If provided, must click into specific Operation row
        """
        print(f"Executing workflow: {action_type} {phase} for order {order_id} (Conf ID: {confirmation_id})...")
        
        self._start_driver_if_not_running()
        
        phase_map = {"setup": ["Setup", "Rüsten"], "processing": ["Processing", "Bearbeitung", "Order", "Auftrag"], "teardown": ["Teardown", "Abrüsten"]}
        action_map = {"start": ["Start", "Beginnen"], "end": ["Finish", "End", "Beenden", "Abschließen"]}
        
        target_text = f"{action_type.capitalize()} {phase}"
        if action_type == "end": target_text = f"Finish {phase}"
        
        def clear_popups():
            try:
                ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(0.3)
                close_btns = self.driver.find_elements(By.XPATH, "//button[@title='Close' or @title='Schließen' or contains(@id, 'message-close')]")
                for btn in close_btns:
                    if btn.is_displayed(): btn.click()
            except: pass

        def try_click_button(xpath_locator):
            try:
                btn = WebDriverWait(self.driver, 2).until(EC.element_to_be_clickable((By.XPATH, xpath_locator)))
                try: ActionChains(self.driver).move_to_element(btn).click().perform(); time.sleep(0.5)
                except: pass
                self.driver.execute_script("arguments[0].click();", btn)
                return True
            except: return False

        # --- 1. Logic Fork ---
        # If no confirmation_id is specified, we can lazily use "Quick Check"
        # If ID is specified, force re-navigation to ensure we enter the correct row
        
        if confirmation_id is None:
            print("🚀 No Confirmation ID specified, attempting quick click...")
            clear_popups()
            xpath = f"//button[.//text()='{target_text}']"
            if try_click_button(xpath):
                print("✅ Quick click successful!")
                time.sleep(3)
                return {"status": "success", "message": f"Executed {target_text} (Quick)"}
        else:
            print(f"🔒 ID {confirmation_id} specified, skipping quick check, executing precise navigation...")

        # --- 2. Full Navigation (Passing confirmation_id) ---
        # This step calls open_app -> search order -> [KEY] click the row containing confirmation_id
        nav_result = self.open_app(
            personnel_number, 
            op_plant, 
            order_id, 
            "Capture of production data", 
            confirmation_id=confirmation_id  # <--- Key parameter
        )
        
        if nav_result["status"] == "error": return nav_result

        # --- 3. Click button after entering detail page ---
        time.sleep(3)
        clear_popups()
        
        xpath = f"//button[.//text()='{target_text}']"
        if try_click_button(xpath):
            return {"status": "success", "message": f"Executed {target_text} for Conf {confirmation_id}"}
        
        # Fuzzy match fallback
        keywords_phase = phase_map.get(phase, ["Processing"])
        keywords_action = action_map.get(action_type, ["Start"])
        for act in keywords_action:
            for ph in keywords_phase:
                fuzzy_xpath = f"//button[contains(., '{act}') and contains(., '{ph}')]"
                if try_click_button(fuzzy_xpath):
                    return {"status": "success", "message": f"Executed {act} {ph} (Fuzzy)"}

        return {"status": "error", "message": f"Could not find button for {target_text}"}

    def quit(self):
        if self.driver:
            self.driver.quit()
            self.driver = None


print("✅ SapRobot Ready.")

# %%
##################################################################
# Cell 3 / Main Script: Flask "Robot Server" (Final)
##################################################################
print("Running Robot Server Cell...")

from flask import Flask, request, jsonify
import json
import threading 

try:
    SapRobot 
except NameError:
    print("!! Critical Error: 'SapRobot' is not defined. Run Cell 1 first.")
    raise

print("Initializing Flask app...")
app = Flask(__name__)
robot = SapRobot() 
print("Robot ready.")

def dispatch_action(data_wrapper: dict):
    print(f"Received data wrapper: {data_wrapper}")
    try:
        if isinstance(data_wrapper, dict) and "action" in data_wrapper:
            command_json = data_wrapper
        else:
            output_string = data_wrapper.get("output", "")
            if not output_string: return
            s = output_string.find('{')
            e = output_string.rfind('}')
            if s == -1 or e == -1: return
            command_json = json.loads(output_string[s:e+1])

        action = command_json.get("action")
        params = command_json.get("params", {}) 

        print(f"▶️ Dispatching: {action} | Params: {params}")
        result = None
        
        if action == "login":
            result = robot.login(**params) 
            
        elif action == "open_app":
            safe = {k: params.get(k) for k in ["personnel_number", "op_plant", "order_id", "app_name", "confirmation_id"]}
            safe = {k: v for k, v in safe.items() if v}
            result = robot.open_app(**safe)
            
        # ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
        # 🟢 Key modification here: Support confirmation_id
        elif action in ["start_order", "end_order"]:
            p_type = "start" if action == "start_order" else "end"
            result = robot.process_order_step(
                personnel_number=params.get("personnel_number", "1"),
                op_plant=params.get("op_plant", "1010"),
                order_id=params.get("order_id"),
                phase=params.get("phase", "processing"),
                action_type=p_type,
                confirmation_id=params.get("confirmation_id") # <--- New
            )
        # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

        elif action == "quit":
            result = robot.quit()
            
        else:
            result = {"status": "error", "message": f"Unknown action: {action}"}
        
        print(f"✅ Result: {result}")

    except Exception as e:
        print(f"!! Error in dispatch: {e}")

@app.route('/execute', methods=['POST'])
def execute_command():
    try:
        raw_data = request.data
        json_string = raw_data.decode('utf-8')
        command_data = json.loads(json_string)
        thread = threading.Thread(target=dispatch_action, args=(command_data,))
        thread.start()
        return jsonify({"status": "received"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == "__main__":
    app.run(host='localhost', port=5000, debug=False, use_reloader=False)


