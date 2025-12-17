# %%
##################################################################
# Cell 1/3: Action Layer (The "Robot") - (Adapt Filters Auto-Fix)
##################################################################
print("Defining Cell 1: SapRobot Action Layer (Adapt Filters Auto-Fix)...")

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains # 引入动作链
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, WebDriverException

# ▼▼▼ 全局变量：防止重启 Kernel 后浏览器丢失，实现“热更新” ▼▼▼
if 'SHARED_DRIVER' not in globals():
    SHARED_DRIVER = None
# ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

class SapRobot:
    def __init__(self):
        global SHARED_DRIVER
        self.driver = None
        self.wait = None
        self.long_wait = None
        
        # 尝试接管已存在的浏览器，不用重新登录
        if SHARED_DRIVER is not None:
            try:
                # 测试连接是否存活
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
        chrome_options.add_experimental_option("detach", True) # 脚本跑完不关浏览器
        chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self._init_waits()
        
        # 保存到全局变量
        SHARED_DRIVER = self.driver
        print("Browser started and saved to global session.")

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
        
# ▼▼▼ 核心逻辑 ▼▼▼
    def open_app(self, personnel_number, op_plant, order_id, app_name, confirmation_id=None):
        print("Executing workflow: open_app (Tab Strategy)...")
        try:
            self._start_driver_if_not_running()
            
            # --- 阶段 1: 智能导航 ---
            try:
                # 检查是否已在 App 内部 (通过 Search 框检查)
                quick_check = WebDriverWait(self.driver, 3)
                quick_check.until(EC.presence_of_element_located(
                    (By.XPATH, "//input[@placeholder='Search']")
                ))
                print("✅ 检测到已在 App 内部 (Search 框可见)。")
            except:
                # 检查工号界面
                on_personnel_screen = False
                try:
                    WebDriverWait(self.driver, 2).until(EC.presence_of_element_located(
                        (By.XPATH, "//input[contains(@id, 'inputPerNr')]")
                    ))
                    on_personnel_screen = True
                    print("✅ 检测到在工号界面...")
                except:
                    pass

                if not on_personnel_screen:
                    print("Step 0: 前往主页...")
                    home_url = 'https://aimprd.advapp.de/sap/bc/ui2/flp#Shell-home'
                    if "Shell-home" not in self.driver.current_url:
                        self.driver.get(home_url)
                    time.sleep(3) 

                    print("Step 1: 点击磁贴...")
                    try:
                        self.long_wait.until(EC.element_to_be_clickable((By.ID, "__tile1-title-inner"))).click()
                    except:
                        self.driver.find_element(By.CSS_SELECTOR, "[title='Advanced Manufacturing']").click()
                    time.sleep(5) 

            # --- 阶段 2: 输工号 ---
            print("Step 2: 检查工号输入...")
            try:
                pn_input = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable(
                    (By.XPATH, "//input[contains(@id, 'inputPerNr')]")
                ))
                pn_input.clear()
                pn_input.send_keys(personnel_number)
                self.driver.find_element(By.XPATH, "//*[text()='Submit']").click()
                print("工号已提交。")
                WebDriverWait(self.driver, 15).until(EC.presence_of_element_located(
                    (By.XPATH, "//input[@placeholder='Search']")
                ))
            except:
                print("无需输工号，继续...")

            # --- 阶段 3: 过滤器处理 ---
            print("Step 3: 处理过滤器...")
            
            # 3.1 检查是否需要 Adapt Filters
            # 我们直接用最笨的方法：如果在屏幕上找不到 'Op. Plant' 这几个字，就去 Adapt
            # 注意：这里我们找的是 Label (文字)，不是输入框，文字通常是有的
            page_source = self.driver.page_source
            if "Op. Plant" not in page_source:
                print("❌ 页面未发现 'Op. Plant' 文字，执行 Adapt Filters...")
                
                print("点击 'Adapt Filters'...")
                self.driver.find_element(By.XPATH, "//*[contains(text(), 'Adapt Filters')]").click()
                time.sleep(2)
                
                print("搜索并勾选...")
                search_box = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='search']")))
                search_box.clear()
                search_box.send_keys("Op. Plant")
                time.sleep(2)
                
                # 强制点击方框
                try:
                    self.driver.find_element(By.XPATH, "//*[text()='Op. Plant']/ancestor::*[contains(@class, 'ListItem') or self::tr]//div[@role='checkbox']").click()
                except:
                    self.driver.find_element(By.XPATH, "//div[@role='checkbox']").click()
                
                self.driver.find_element(By.XPATH, "//*[text()='OK']").click()
                time.sleep(3)
            else:
                print("✅ 页面已包含 Op. Plant 文字，跳过设置。")

            # --- Step 3.3 填数据 (Tab 键大法) ---
            print("Step 3.3: 使用键盘导航填数据...")
            
            # 1. 找到“锚点”：Search 框
            search_input = self.wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//input[@placeholder='Search']")
            ))
            search_input.click()
            print("✅ 已定位 Search 框")
            time.sleep(0.5)
            
            # 2. Tab 到 Order (通常是 Search 的下一个)
            # 使用 ActionChains 模拟按 Tab
            actions = ActionChains(self.driver)
            
            print(f"👉 Tab 到 Order，输入: {order_id}")
            actions.send_keys(Keys.TAB).perform()
            time.sleep(0.5)
            
            # 这里的 active_element 就是当前光标所在的框，不管它有没有 ID
            self.driver.switch_to.active_element.send_keys(order_id)
            time.sleep(0.5)

            # 3. Tab 到 Op. Plant (通常是 Order 的下一个)
            print(f"👉 Tab 到 Op. Plant，输入: {op_plant}")
            actions.send_keys(Keys.TAB).perform()
            time.sleep(0.5)
            
            self.driver.switch_to.active_element.send_keys(op_plant)
            self.driver.switch_to.active_element.send_keys(Keys.TAB) # 再按一下确认
            time.sleep(1)

            # 4. 点击 Go
            print("Step 4: 点击 Go...")
            try:
                self.driver.find_element(By.XPATH, "//button[.='Go']").click()
            except:
                self.driver.find_element(By.XPATH, "//*[text()='Go']").click()
            time.sleep(3)

            # 5. 点击订单
            print(f"Step 5: 点击订单 {order_id}...")
            order_row = self.long_wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"//*[text()='{order_id}']")
            ))
            order_row.click()
            time.sleep(2)

            # 6. 打开最终 App
            print(f"Step 6: 打开 App '{app_name}'...")
            try:
                self.long_wait.until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, f"[title='{app_name}']")
                )).click()
            except:
                print("Title匹配失败，尝试文字匹配...")
                short_name = app_name.split(' ')[0]
                self.driver.find_element(By.XPATH, f"//*[contains(text(), '{short_name}')]").click()

            print(f"成功点击 '{app_name}' Tile!")
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

print("✅ Cell 1 (Tab Strategy) Ready.")

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


