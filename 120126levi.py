# %%
##################################################################
# Cell 1/3: Action Layer (Bilingual App Tile Fix)
##################################################################

print("Defining SapRobot Action Layer (Bilingual App Tile Fix)...")

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

if "SHARED_DRIVER" not in globals():
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
            except Exception:
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
            login_url = "https://aimprd.advapp.de/sap/bc/ui2/flp"
            self.driver.get(login_url)

            user_field = self.wait.until(
                EC.presence_of_element_located((By.ID, "USERNAME_FIELD-inner"))
            )
            pass_field = self.driver.find_element(By.ID, "PASSWORD_FIELD-inner")

            user_field.send_keys(username)
            pass_field.send_keys(password)
            self.driver.find_element(By.ID, "LOGIN_LINK").click()

            self.long_wait.until(
                EC.presence_of_element_located((By.ID, "shell-header-logo"))
            )
            print(f"Successfully logged in as {username}.")
            return {"status": "success", "message": "Login successful"}

        except Exception as e:
            return {"status": "error", "message": str(e)}

    # ▼▼▼ Core Navigation (Bilingual + App Translation) ▼▼▼
    def open_app(self, personnel_number, operation_filter, order_id, app_name):
        print(
            f"Executing workflow: open_app for '{app_name}' (Filter Op: {operation_filter})..."
        )
        try:
            self._start_driver_if_not_running()

            # --- Phase 1: Smart Navigation (Search/Suchen) ---
            try:
                quick_check = WebDriverWait(self.driver, 3)
                quick_check.until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//input[@placeholder='Search' or @placeholder='Suchen']")
                    )
                )
                print("✅ Already inside app (Search/Suchen visible).")
            except Exception:
                on_personnel_screen = False
                try:
                    WebDriverWait(self.driver, 2).until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//input[contains(@id, 'inputPerNr')]")
                        )
                    )
                    on_personnel_screen = True
                except Exception:
                    pass

                if not on_personnel_screen:
                    print("Step 0: Go to home...")
                    home_url = "https://aimprd.advapp.de/sap/bc/ui2/flp#Shell-home"
                    if "Shell-home" not in self.driver.current_url:
                        self.driver.get(home_url)
                    time.sleep(3)

                    print("Step 1: Click tile...")
                    try:
                        self.long_wait.until(
                            EC.element_to_be_clickable((By.ID, "__tile1-title-inner"))
                        ).click()
                    except Exception:
                        self.driver.find_element(
                            By.CSS_SELECTOR, "[title='Advanced Manufacturing']"
                        ).click()
                    time.sleep(5)

            # --- Phase 2: Personnel Number ---
            print("Step 2: Check personnel input...")
            try:
                pn_input = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//input[contains(@id, 'inputPerNr')]")
                    )
                )
                pn_input.clear()
                pn_input.send_keys(personnel_number)
                try:
                    self.driver.find_element(By.XPATH, "//*[text()='Senden']").click()
                except Exception:
                    try:
                        self.driver.find_element(By.XPATH, "//*[text()='Submit']").click()
                    except Exception:
                        pn_input.send_keys(Keys.ENTER)

                WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//input[@placeholder='Search' or @placeholder='Suchen']")
                    )
                )
            except Exception:
                print("No personnel input needed, continuing...")

            # --- Phase 3: Filters (EN: Operation / DE: Vorgang) ---
            print("Step 3: Handle filters (Operation/Vorgang Check)...")
            operation_field_visible = False

            try:
                self.driver.find_element(
                    By.XPATH,
                    "//input[contains(@aria-label, 'Operation') or contains(@placeholder, 'Operation') "
                    "or contains(@aria-label, 'Vorgang') or contains(@placeholder, 'Vorgang')]",
                )
                operation_field_visible = True
            except Exception:
                try:
                    self.driver.find_element(
                        By.XPATH,
                        "//*[text()='Operation' or text()='Vorgang']/following::input[1]",
                    )
                    operation_field_visible = True
                except Exception:
                    pass

            if not operation_field_visible:
                print("❌ 'Operation/Vorgang' filter not found -> Adapt Filters...")
                self.driver.find_element(
                    By.XPATH,
                    "//*[contains(text(), 'Adapt Filters') or contains(text(), 'Filter anpassen')]",
                ).click()
                time.sleep(2)

                search_box = self.wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//input[@type='search' or @placeholder='Search' or @placeholder='Suchen' "
                            "or @placeholder='Nach Filtern suchen']",
                        )
                    )
                )
                search_box.clear()

                is_german = (
                    "Suchen" in self.driver.page_source
                    or "Filter anpassen" in self.driver.page_source
                )
                search_term = "Vorgang" if is_german else "Operation"

                print(f"Searching for filter: {search_term}")
                search_box.send_keys(search_term)
                time.sleep(2)

                print(f"Selecting '{search_term}' checkbox...")
                try:
                    target_row_xpath = (
                        f"//tr[.//text()[normalize-space()='{search_term}']][.//div[@role='checkbox']]"
                    )
                    checkbox = self.driver.find_element(
                        By.XPATH, target_row_xpath + "//div[@role='checkbox']"
                    )
                    checkbox.click()
                    print(f"✅ Clicked '{search_term}' checkbox")
                except Exception:
                    print("⚠️ Precise selection failed, trying loop method...")
                    try:
                        rows = self.driver.find_elements(
                            By.XPATH, "//tr[.//div[@role='checkbox']]"
                        )
                        for row in rows:
                            if row.text.strip() == search_term:
                                row.find_element(By.XPATH, ".//div[@role='checkbox']").click()
                                break
                    except Exception:
                        pass

                self.driver.find_element(By.XPATH, "//*[text()='OK']").click()
                time.sleep(3)
            else:
                print("✅ 'Operation/Vorgang' filter present, skip Adapt.")

            # --- Phase 3.3: Tab Input Strategy ---
            print("Step 3.3: Keyboard navigation input...")

            search_input = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//input[@placeholder='Search' or @placeholder='Suchen']")
                )
            )
            search_input.click()
            time.sleep(0.2)
            search_input.send_keys(Keys.CONTROL, "a")
            search_input.send_keys(Keys.DELETE)
            time.sleep(0.2)

            actions = ActionChains(self.driver)

            # Tab to Order / Auftrag
            print(f"👉 Tab to Order/Auftrag, input: {order_id}")
            actions.send_keys(Keys.TAB).perform()
            time.sleep(0.4)
            self.driver.switch_to.active_element.send_keys(order_id)

            # Tab to Operation / Vorgang
            if operation_filter:
                print(f"👉 Tab to Operation/Vorgang, input: {operation_filter}")
                actions.send_keys(Keys.TAB).perform()
                time.sleep(0.4)
                self.driver.switch_to.active_element.send_keys(operation_filter)

            print("Step 4: Click Go/Start...")
            try:
                self.driver.find_element(By.XPATH, "//button[.='Start']").click()
            except Exception:
                try:
                    self.driver.find_element(By.XPATH, "//button[.='Go']").click()
                except Exception:
                    self.driver.find_element(
                        By.XPATH, "//*[text()='Start' or text()='Go']"
                    ).click()
            time.sleep(3)

            # 5. Click Order Row
            print(f"Step 5: Click Order {order_id}...")
            row = self.long_wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, f"//*[text()='{order_id}']/ancestor::tr[1]")
                )
            )
            row.click()
            time.sleep(2)

            # --- 6. Open Final App (auto translate app name) ---
            print(f"Step 6: Open App '{app_name}'...")

            app_translations = {
                "Capture of production data": "Erfassung von Produktionsdaten",
                "Business Documents": "Business Documents",
                "Production Monitor": "Produktionsmonitor",
            }

            de_name = app_translations.get(app_name, app_name)

            print(f"   -> Looking for EN: '{app_name}' OR DE: '{de_name}'")

            xpath_locator = (
                f"//*[@title='{app_name}' or @title='{de_name}' "
                f"or contains(text(), '{app_name}') or contains(text(), '{de_name}')]"
            )

            try:
                self.long_wait.until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, f"[title='{app_name}'], [title='{de_name}']")
                    )
                ).click()
            except Exception:
                try:
                    self.driver.find_element(By.XPATH, xpath_locator).click()
                except Exception as e:
                    print(f"❌ Failed to open app tile via Title or Text. Error: {e}")
                    short_name = app_name.split(" ")[0]
                    self.driver.find_element(
                        By.XPATH, f"//*[contains(text(), '{short_name}')]"
                    ).click()

            print(f"Clicked '{app_name}' / '{de_name}' Tile!")
            return {"status": "success", "message": f"Opened: {app_name}"}

        except Exception as e:
            print(f"!! Workflow failed: {e}")
            return {"status": "error", "message": str(e)}

    # ▼▼▼ Intelligent Processing (Bilingual: EN/DE) ▼▼▼
    def process_order_step(self, personnel_number, operation_filter, order_id, phase, action_type):
        print(
            f"Executing: {action_type} {phase} | Order: {order_id} | OpFilter: {operation_filter}"
        )
        self._start_driver_if_not_running()

        phase_map = {
            "setup": ["Setup", "Rüsten"],
            "processing": ["Processing", "Bearbeitung", "Bearbeiten"],
            "teardown": ["Teardown", "Abrüsten"],
        }

        action_map = {
            "start": ["Start", "starten", "Starten"],
            "end": ["Finish", "End", "beenden", "Beenden", "abschließen"],
        }

        possible_texts = []
        for v in action_map.get(action_type, []):
            for p in phase_map.get(phase, []):
                possible_texts.append(f"{v} {p}")
                possible_texts.append(f"{p} {v}")

        print(f"🔍 Looking for buttons: {possible_texts}")

        def try_click_button(xpath_locator):
            try:
                btn = WebDriverWait(self.driver, 1.0).until(
                    EC.element_to_be_clickable((By.XPATH, xpath_locator))
                )
                print(f"⚡️ Found button '{xpath_locator}', clicking...")
                self.driver.execute_script("arguments[0].style.border='3px solid red'", btn)
                time.sleep(0.5)

                try:
                    ActionChains(self.driver).move_to_element(btn).click().perform()
                    return True
                except Exception:
                    pass
                try:
                    btn.click()
                    return True
                except Exception:
                    pass
                self.driver.execute_script("arguments[0].click();", btn)
                return True
            except Exception:
                return False

        def clear_popups():
            try:
                ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                close_btns = self.driver.find_elements(
                    By.XPATH, "//button[@title='Close' or @title='Schließen']"
                )
                for btn in close_btns:
                    if btn.is_displayed():
                        btn.click()
            except Exception:
                pass

        looks_like_correct_page = False
        if operation_filter:
            try:
                if (
                    len(
                        self.driver.find_elements(
                            By.XPATH,
                            f"//*[(contains(text(), 'Operation') or contains(text(), 'Vorgang')) "
                            f"and contains(text(), '{operation_filter}')]",
                        )
                    )
                    > 0
                ):
                    print(f"✅ Screen matches Op/Vorgang {operation_filter}.")
                    looks_like_correct_page = True
            except Exception:
                pass

        if not looks_like_correct_page and order_id:
            try:
                if order_id in self.driver.page_source:
                    print(f"✅ Screen matches Order/Auftrag {order_id}.")
                    looks_like_correct_page = True
            except Exception:
                pass

        if looks_like_correct_page:
            clear_popups()
            for txt in possible_texts:
                if try_click_button(f"//button[.//text()='{txt}']"):
                    return {"status": "success", "message": f"Executed {txt} (Instant)"}

            if action_type == "end":
                print("⚠️ Attempting other Finish/Beenden buttons...")
                for p_key in ["setup", "processing", "teardown"]:
                    for v in action_map["end"]:
                        for p_val in phase_map[p_key]:
                            for t in [f"{v} {p_val}", f"{p_val} {v}"]:
                                if try_click_button(f"//button[.//text()='{t}']"):
                                    return {"status": "success", "message": f"Executed {t} (Auto-Fix)"}

            print("❌ Button not found, preparing to refresh...")

        print("🔄 Executing full navigation flow...")
        nav_result = self.open_app(
            personnel_number,
            operation_filter,
            order_id,
            "Capture of production data",
        )
        if nav_result["status"] == "error":
            return nav_result

        time.sleep(3)
        clear_popups()

        for txt in possible_texts:
            if try_click_button(f"//button[.//text()='{txt}']"):
                return {"status": "success", "message": f"Executed {txt} after nav"}

        return {"status": "error", "message": f"Could not find button. Tried: {possible_texts}"}

    # ▼▼▼ Next/Previous/Quit (Bilingual) ▼▼▼
    def next_order(self):
        print("Executing workflow: next_order...")
        strategies = [
            (By.XPATH, "//button[.//text()='Next']"),
            (By.XPATH, "//button[.//text()='Nächster']"),
            (By.XPATH, "//button[.//bdi[contains(text(), 'Next')]]"),
            (By.XPATH, "//button[.//bdi[contains(text(), 'Nächster')]]"),
            (By.CSS_SELECTOR, "button[title='Next']"),
            (By.CSS_SELECTOR, "button[title='Nächster']"),
            (By.XPATH, "//span[contains(@class, 'sapUiIcon') and @data-sap-ui-icon-content='']"),
        ]
        return self._click_nav_button(strategies, "Next/Nächster")

    def previous_order(self):
        print("Executing workflow: previous_order...")
        strategies = [
            (By.XPATH, "//button[.//text()='Previous']"),
            (By.XPATH, "//button[.//text()='Vorheriger']"),
            (By.XPATH, "//button[.//bdi[contains(text(), 'Previous')]]"),
            (By.XPATH, "//button[.//bdi[contains(text(), 'Vorheriger')]]"),
            (By.CSS_SELECTOR, "button[title='Previous']"),
            (By.CSS_SELECTOR, "button[title='Vorheriger']"),
            (By.XPATH, "//span[contains(@class, 'sapUiIcon') and @data-sap-ui-icon-content='']"),
        ]
        return self._click_nav_button(strategies, "Previous/Vorheriger")

    def _click_nav_button(self, strategies, name):
        self._start_driver_if_not_running()
        print(f"🔍 Starting search for {name} button...")
        try:
            ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
        except Exception:
            pass

        for by, val in strategies:
            try:
                elements = self.driver.find_elements(by, val)
                target_btn = None
                for btn in elements:
                    if btn.is_displayed():
                        target_btn = btn
                        break

                if target_btn:
                    print(f"✅ Found {name}: {val}")
                    self.driver.execute_script("arguments[0].style.border='3px solid red'", target_btn)
                    time.sleep(0.5)
                    try:
                        ActionChains(self.driver).move_to_element(target_btn).click().perform()
                        return {"status": "success", "message": f"Clicked {name}"}
                    except Exception:
                        pass
                    try:
                        target_btn.click()
                        return {"status": "success", "message": f"Clicked {name}"}
                    except Exception:
                        pass
                    self.driver.execute_script("arguments[0].click();", target_btn)
                    time.sleep(3)
                    return {"status": "success", "message": f"Clicked {name}"}
            except Exception:
                continue

        return {"status": "error", "message": f"Not found {name}"}

    def quit(self):
        if self.driver:
            self.driver.quit()
            self.driver = None


print("✅ SapRobot Ready (Bilingual App Tile Fix).")

# %%
##################################################################
# Cell 3 / Main Script: Flask Server (Clean Version)
##################################################################
print("Running Robot Server Cell (Clean Version)...")

from flask import Flask, request, jsonify
import json
import threading
import traceback

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

    command_json = None
    try:
        if isinstance(data_wrapper, dict) and "action" in data_wrapper:
            command_json = data_wrapper
        else:
            output_string = data_wrapper.get("output", "")
            if not output_string:
                return

            s = output_string.find("{")
            e = output_string.rfind("}")
            if s == -1 or e == -1:
                return

            json_str = output_string[s : e + 1]
            print(f"🔍 Extracting: {json_str}")

            try:
                command_json = json.loads(json_str)
            except json.JSONDecodeError:
                print("⚠️ JSON Error, applying auto-fix...")
                try:
                    command_json = json.loads(json_str + "}")
                except Exception:
                    return

    except Exception:
        return

    try:
        action = command_json.get("action")
        params = command_json.get("params", {})

        print(f"▶️ Dispatching: {action} | Params: {params}")
        result = None

        if action == "login":
            result = robot.login(**params)

        elif action == "open_app":
            safe = {
                "personnel_number": params.get("personnel_number", "1"),
                "operation_filter": params.get("operation_filter", ""),
                "order_id": params.get("order_id"),
                "app_name": params.get("app_name"),
            }
            safe = {k: v for k, v in safe.items() if v is not None}
            result = robot.open_app(**safe)

        elif action in ["start_order", "end_order"]:
            p_type = "start" if action == "start_order" else "end"

            # Wenn es ein Auftrag-Start ist → Phase = setup (Rüsten starten)
            if action == "start_order":
                phase = "setup"
            else:
                phase = params.get("phase", "processing")

            result = robot.process_order_step(
                personnel_number=params.get("personnel_number", "1"),
                operation_filter=params.get("operation_filter", ""),
                order_id=params.get("order_id"),
                phase=phase,
                action_type=p_type,
            )

        elif action == "next_order":
            result = robot.next_order()

        elif action == "previous_order":
            result = robot.previous_order()

        elif action == "quit":
            result = robot.quit()

        else:
            result = {"status": "error", "message": f"Unknown action: {action}"}

        print(f"✅ Result: {result}")

    except Exception as e:
        print(f"!! Error in dispatch: {e}")
        traceback.print_exc()


@app.route("/execute", methods=["POST"])
def execute_command():
    try:
        raw_data = request.data
        json_string = raw_data.decode("utf-8")
        command_data = json.loads(json_string)
        thread = threading.Thread(target=dispatch_action, args=(command_data,))
        thread.start()
        return jsonify({"status": "received"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


if __name__ == "__main__":
    app.run(host="localhost", port=5000, debug=False, use_reloader=False)
