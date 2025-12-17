# %%
##################################################################
# Cell 1/3: Action Layer (The "Robot") - (Final Robust Version)
##################################################################
print("Defining Cell 1: SapRobot Action Layer (Final Robust Version)...")

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, WebDriverException

class SapRobot:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.long_wait = None
        print("Robot is ready (will auto-download driver and use stable options).")



    def _start_driver_if_not_running(self):
        if self.driver is not None:
            try:
                _ = self.driver.window_handles
            except Exception as e:
                print(f"Browser closed/crashed. Restarting...")
                self.driver = None

        if self.driver is None:

            print("Starting a new browser instance...")
            chrome_options = Options()
            chrome_options.add_argument("--start-maximized")
            # 关键：防止代码结束后浏览器关闭
            chrome_options.add_experimental_option("detach", True)
            chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.wait = WebDriverWait(self.driver, 10)
            self.long_wait = WebDriverWait(self.driver, 20)
            print("Browser started.")



    def login(self, username, password):
        # ... (登录代码保持不变，略以节省空间，核心是下面的 open_app) ...
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
    # ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    # (Modified) 智能判断当前位置，支持“断点续传”
    def open_app(self, personnel_number, op_plant, order_id, app_name, confirmation_id=None):
        print("Executing workflow: open_app (Smart Resume)...")
        try:
            self._start_driver_if_not_running()
            # 1. 智能检测：我们已经在“人事编号”界面了吗？
            # -------------------------------------------------------
            print("Checking current screen state...")
            already_on_personnel_screen = False
            try:
                # 快速检查一下当前页面有没有“Personnel number”输入框
                quick_check = WebDriverWait(self.driver, 2)
                quick_check.until(EC.presence_of_element_located(
                    (By.XPATH, "//input[contains(@aria-label, 'Personnel number')]")
                ))
                print("✅ Detected: Already on Personnel Number screen. Skipping navigation.")
                already_on_personnel_screen = True
            except:
                print("Not on Personnel screen. Will perform full navigation.")
            # 2. 如果不在人事界面，才执行“回主页 -> 点磁贴”
            # -------------------------------------------------------
            if not already_on_personnel_screen:
                print("Step 0: Navigating to Home Page...")
                home_url = 'https://aimprd.advapp.de/sap/bc/ui2/flp#Shell-home'
                # 只有当 URL 真的不同时才跳转，避免不必要的刷新
                if "Shell-home" not in self.driver.current_url:
                    self.driver.get(home_url)
                time.sleep(3) # 等待主页加载
                print("Step 1: Clicking 'Advanced Manufacturing' tile...")
                # 使用最简单的文字匹配点击磁贴
                am_tile = self.long_wait.until(EC.element_to_be_clickable(
                    (By.ID, "__tile1-title-inner")
                ))
                am_tile.click()
                time.sleep(2)

            # 3. 输入人事编号 (无论是否跳过导航，这一步都要做)
            # -------------------------------------------------------
            print("Step 2: Handling Personnel Number...")
            try:
                # 等待输入框出现
                popup_wait = WebDriverWait(self.driver, 5)
                pn_input = popup_wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//input[contains(@id, 'inputPerNr')]")
                ))

                # 只有当输入框是空的，或者我们需要确信它是正确的时候才输入
                pn_input.clear()
                pn_input.send_keys(personnel_number)
                # 点击 Submit
                self.driver.find_element(By.XPATH, "//*[text()='Submit']").click()
                print(f"Personnel number {personnel_number} submitted.")
                time.sleep(3)

            except TimeoutException:
                # 如果找不到输入框，可能已经登录进去了，或者是那个“设备名”弹窗挡住了
                print("⚠️ Warning: Personnel number input not found. Checking if we are already logged in...")
            # 4. 输入过滤器 (Op. Plant, Order, etc.)
            # -------------------------------------------------------
            print("Step 3: Waiting for filter bar...")
            # 这里增加容错，等待 Op. Plant 输入框
            filter_op_plant = self.long_wait.until(EC.presence_of_element_located(
                (By.XPATH, "//input[contains(@aria-label, 'Op. Plant')]")
            ))
            print(f"Entering Op. Plant: {op_plant}")
            filter_op_plant.clear()
            filter_op_plant.send_keys(op_plant)
            print(f"Entering Order: {order_id}")
            order_input = self.driver.find_element(By.XPATH, "//input[contains(@aria-label, 'Order')]")
            order_input.clear()
            order_input.send_keys(order_id)
            if confirmation_id:
                print(f"Entering Confirmation: {confirmation_id}")
                conf_input = self.driver.find_element(By.XPATH, "//input[contains(@aria-label, 'Confirmation')]")
                conf_input.clear()
                conf_input.send_keys(confirmation_id)

            # 5. 点击 Go 并选择订单
            # -------------------------------------------------------
            print("Step 4: Clicking 'Go'...")
            self.driver.find_element(By.XPATH, "//*[text()='Go']").click()
            time.sleep(3)
            print(f"Step 5: Waiting for order '{order_id}'...")
            # 查找包含订单号的文本并点击
            order_row = self.long_wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"//*[text()='{order_id}']")
            ))
            order_row.click()
            print(f"Clicked on order {order_id}.")

            # 6. 打开最终 App
            # -------------------------------------------------------
            print(f"Step 6: Waiting for '{app_name}' tile...")
            app_tile = self.long_wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"//*[text()='{app_name}']")
            ))
            app_tile.click()
            print(f"Clicked '{app_name}' tile.")
            return {"status": "success", "message": f"Workflow complete. Clicked app: {app_name}"}

        except Exception as e:
            print(f"!! Workflow 'open_app' failed: {e}")
            # 打印详细错误，但不退出
            import traceback
            traceback.print_exc()
            return {"status": "error", "message": str(e)}
    # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    def quit(self):
        # 只有明确调用 quit 时才关闭
        if self.driver is not None:
            try:
                self.driver.quit()
            except:
                pass
            finally:
                self.driver = None
                self.wait = None
                self.long_wait = None
            return {"status": "success", "message": "Browser closed."}
        else:
            return {"status": "success", "message": "Browser was not running."}

print("✅ Cell 1/3: SapRobot Action Layer (Final Robust) defined.")##################################################################

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
        # Fall A: n8n schickt direkt {action, params, ...}
        if "action" in data_wrapper:
            command_json = data_wrapper

        # Fall B: n8n schickt {"output": "...{...}..."} (dein bisheriges Format)
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
        params = command_json.get("params", {}) 

        print(f"Dispatching action: {action} with params: {params}")

        result = None
        if action == "login":
            result = robot.login(**params) 
            
        # ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
        # (New) 链接到您的 "apps öffnen" 合并动作
        elif action == "open_app":
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
            
            result = robot.open_app(**safe_params)
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