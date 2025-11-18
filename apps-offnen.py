# %%
##################################################################
# Cell 1/3: Action Layer (The "Robot") - (Final Robust Version)
##################################################################
print("Defining Cell 1: SapRobot Action Layer (Final Robust Version)...")

# --- 1. Import Selenium and Webdriver-Manager ---
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options 
from selenium.common.exceptions import TimeoutException
class SapRobot:
    def __init__(self): 
        self.driver = None
        self.wait = None
        self.long_wait = None 
        print("Robot is ready (will auto-download driver and use stable options).")

    def _start_driver_if_not_running(self):
        if self.driver is None:
            print("Starting a new browser instance...")
            chrome_options = Options()
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_argument("--remote-debugging-port=9222")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
            print("Webdriver-Manager is checking/downloading driver...")
            service = Service(ChromeDriverManager().install())
            print("Driver check complete.")
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.wait = WebDriverWait(self.driver, 10) 
            self.long_wait = WebDriverWait(self.driver, 20) 
            print("Browser started.")
        else:
            print("Browser is already running.")

    def login(self, username, password):
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
            self.long_wait.until(
                EC.presence_of_element_located((By.ID, 'shell-header-logo')) 
            )
            print(f"Successfully logged in as {username}.")
            return {"status": "success", "message": "Login successful"}
        
        except Exception as e:
            print(f"Login failed: {e}")
            self.quit() 
            print("Robot state has been reset.")
            return {"status": "error", "message": str(e)}

    # ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    # (New) 这是您的 "apps öffnen" 动作 (合并版)
    # 它现在接受 'app_name' 并且 confirmation_id 是可选的
    def select_order_and_open_app(self, personnel_number, op_plant, order_id, app_name, confirmation_id=None):
        """
        Action: 自动化从 Fiori 启动板到为特定订单打开 *指定应用磁贴* 的完整流程。
        (此版本包含人事编号逻辑)
        
        Steps:
        1. 点击 'Advanced Manufacturing' 磁贴。
        2. 输入人事编号 (personnel_number) 并提交。
        3. 输入过滤器 (Op. Plant, Order, Confirmation(可选))。
        4. 点击 'Go' 搜索。
        5. 点击列表中的结果订单。
        6. 点击 'app_name' 参数指定的应用磁贴 (例如 "Business Documents" 或 "PQ1")。
        """
        print("Executing workflow: select_order_and_open_app (Combined Version)...")
        try:
            if self.driver is None:
                return {"status": "error", "message": "Driver not running. Please login first."}

            # --- 步骤 1 & 2: 点击磁贴并输入人事编号 ---
            print("Step 1: Clicking 'Advanced Manufacturing' tile...")
            am_tile = self.long_wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(@class, 'sapMGT') and .//span[text()='Advanced Manufacturing']]")))
            
            am_tile.click()
            print("Step 2: Entering personnel number...")
            
            try:
                popup_wait = WebDriverWait(self.driver, 5) # 用 5s 等待弹窗
                pn_input = popup_wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//input[contains(@aria-label, 'Personnel number')]")
                ))
                pn_input.send_keys(personnel_number)
                
                self.driver.find_element(By.XPATH, "//button[.//span[text()='Submit']]").click()
                print(f"Personnel number {personnel_number} submitted.")
            except TimeoutException:
                print("Personnel number popup not found, assuming already submitted.")

            # --- 步骤 3 & 4: 输入过滤器并点击 Go ---
            print("Step 3: Waiting for filter bar and entering filters...")
            
            filter_op_plant = self.long_wait.until(EC.presence_of_element_located(
                (By.XPATH, "//input[contains(@aria-label, 'Op. Plant')]")
            ))
            
            print(f"Entering Op. Plant: {op_plant}")
            filter_op_plant.send_keys(op_plant)
            
            print(f"Entering Order: {order_id}")
            self.driver.find_element(By.XPATH, "//input[contains(@aria-label, 'Order')]").send_keys(order_id)
            
            # (New) 只有在提供了 confirmation_id 时才输入
            if confirmation_id:
                print(f"Entering Confirmation: {confirmation_id}")
                self.driver.find_element(By.XPATH, "//input[contains(@aria-label, 'Confirmation')]").send_keys(confirmation_id)
            else:
                print("Confirmation ID not provided, skipping.")


            print("Step 4: Clicking 'Go'...")
            self.driver.find_element(By.XPATH, "//button[.//span[text()='Go']]").click()

            # --- 步骤 5: 从列表中选择订单 ---
            print(f"Step 5: Waiting for order '{order_id}' in results...")
            
            order_row = self.long_wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"//table//td//a[contains(text(), '{order_id}')]")
            ))
            order_row.click()
            print(f"Clicked on order {order_id}.")

            # --- 步骤 6: 打开指定的 App 磁贴 ---
            print(f"Step 6: Waiting for '{app_name}' tile...")
            
            app_tile = self.long_wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"//*[contains(@aria-label, '{app_name}')]")
            ))
            app_tile.click()
            print(f"Clicked '{app_name}' tile.")
            
            time.sleep(2) 

            print("Workflow complete.")
            return {"status": "success", "message": f"Workflow complete. Clicked app: {app_name}"}

        except Exception as e:
            print(f"!! Workflow 'select_order_and_open_app' failed: {e}")
            return {"status": "error", "message": str(e)}
    # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲


    def quit(self):
        """Action: Close the browser and end the session."""
        if self.driver is not None:
            print("Quit command received. Closing browser...")
            try:
                self.driver.quit()
                print("Browser closed.")
            except Exception as e:
                print(f"Error while quitting: {e}")
            finally:
                self.driver = None
                self.wait = None
                self.long_wait = None
            return {"status": "success", "message": "Browser closed."}
        else:
            print("Quit command received, but browser is not running.")
            return {"status": "success", "message": "Browser was not running."}

print("✅ Cell 1/3: SapRobot Action Layer (Final Robust) defined.")

# %%
##################################################################
# Cell 3 / Main Script: Flask "Robot Server" (Final Robust Version)
# Run this cell to start the robot "listener" service
##################################################################
print("Running Robot Server Cell...")

# --- 1. Import necessary libraries ---
from flask import Flask, request, jsonify
import json
import threading 

# --- 2. Ensure SapRobot from Cell 1 is defined ---
try:
    SapRobot 
except NameError:
    print("="*50); print("!! Critical Error: 'SapRobot' is not defined !!"); print("!! You must run Cell 1 first !!"); print("="*50)
    raise

# --- 3. Initialize Flask app and Robot instance ---
print("Initializing Flask app and SapRobot...")
app = Flask(__name__)
robot = SapRobot() 
print("Robot is ready and waiting for commands.")

# --- 4. Define the "dispatcher" function (now aware of 'quit') ---
def dispatch_action(data_wrapper: dict):
    print(f"Received data wrapper for dispatch: {data_wrapper}")
    try:
        output_string = data_wrapper.get("output")
        if not output_string:
            print("!! Error: 'output' key is missing or empty.")
            return
        
        start_index = output_string.find('{')
        end_index = output_string.rfind('}')
        if start_index == -1 or end_index == -1:
            print(f"!! Error: Could not find JSON {{...}} in string: {output_string}")
            return
            
        json_string_cleaned = output_string[start_index : end_index + 1]
        print(f"Cleaned JSON string: {json_string_cleaned}")
        command_json = json.loads(json_string_cleaned)
        
        action = command_json.get("action")
        params = command_json.get("params", {}) 

        print(f"Dispatching action: {action} with params: {params}")

        result = None
        if action == "login":
            result = robot.login(**params) 
            
        # ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
        # (New) 链接到您的 "apps öffnen" 合并动作
        elif action == "open_order_app":
            # (New) 安全地提取参数
            # 我们使用 .get() 来安全地处理可选的 confirmation_id
            safe_params = {
                "personnel_number": params.get("personnel_number"),
                "op_plant": params.get("op_plant"),
                "order_id": params.get("order_id"),
                "app_name": params.get("app_name"),
                "confirmation_id": params.get("confirmation_id") # 如果不存在，会是 None
            }
            # 移除值为 None 的键，以匹配 Cell 1 中的默认参数
            safe_params = {k: v for k, v in safe_params.items() if v is not None}
            
            result = robot.select_order_and_open_app(**safe_params)
        # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
            
        elif action == "quit":
            result = robot.quit()
            
        elif action == "open_worklist":
            print("Action 'open_worklist' recognized (but not implemented in Cell 1)")
            result = {"status": "pending", "message": "open_worklist not implemented"}
        elif action == "unknown":
            print("Ollama could not understand the command.")
            result = {"status": "error", "message": "Unknown command"}
        else:
            print(f"Error: Received unknown action '{action}'")
            result = {"status": "error", "message": f"Unknown action: {action}"}
        
        print(f"Action '{action}' finished with result: {result}")

    except Exception as e:
        print(f"!! Error during action dispatch: {e}")

# --- 5. Create the Webhook receiver endpoint (unchanged) ---
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

# --- 6. Start the server (unchanged) ---
if __name__ == "__main__":
    print("====================================================")
    print("=== Python Robot Server is starting... ===")
    print("=== Listening on http://localhost:5000/execute ===")
    print("====================================================")
    app.run(host='localhost', port=5000, debug=False, use_reloader=False)


