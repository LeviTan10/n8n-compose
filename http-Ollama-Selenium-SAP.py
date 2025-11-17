# %%
%pip install Flask

# %%
%pip install webdriver-manager

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

# (New) Import webdriver-manager and Chrome Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options 

class SapRobot:
    def __init__(self): 
        self.driver = None
        self.wait = None
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
            self.wait.until(
                EC.presence_of_element_located((By.ID, 'shell-header-logo')) 
            )
            print(f"Successfully logged in as {username}.")
            return {"status": "success", "message": "Login successful"}
        
        # ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
        # (New) Robust error handling and automatic reset
        except Exception as e:
            print(f"Login failed: {e}")
            print("Attempting to quit driver after failure...")
            self.quit() # Call our new quit method to clean up
            print("Robot state has been reset.")
            return {"status": "error", "message": str(e)}
        # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    # ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    # (New) Add a Quit action so you can close the browser correctly
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
                # The most important step: reset the robot's state
                self.driver = None
                self.wait = None
            return {"status": "success", "message": "Browser closed."}
        else:
            print("Quit command received, but browser is not running.")
            return {"status": "success", "message": "Browser was not running."}
    # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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
        # (New) Added handling for the 'quit' action
        elif action == "quit":
            result = robot.quit()
        # ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
            
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


