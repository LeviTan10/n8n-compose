# %%
##################################################################
# Cell 1/3: Action Layer (The "Robot")
# Run this cell to "define" your robot in memory
##################################################################
print("Defining Cell 1: SapRobot Action Layer...")

# --- 1. Import all Selenium tools ---
from selenium import webdriver
from selenium.webdriver.chrome.service import Service  # <-- Import Service (to fix startup issues)
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

class SapRobot:
    def __init__(self, driver_path):
        # On init, don't start the browser yet
        self.driver_path = driver_path
        self.driver = None
        self.wait = None  # Add a wait object
        print("Robot is ready.")

    def _start_driver_if_not_running(self):
        """Internal function: Check if the browser is running, start it if not"""
        if self.driver is None:
            print("Starting a new browser instance...")
            
            # --- This is the correct Selenium 4+ startup method ---
            service = Service(executable_path=self.driver_path)
            self.driver = webdriver.Chrome(service=service)
            self.driver.maximize_window()
            # Create a "smart wait" for the entire robot instance
            self.wait = WebDriverWait(self.driver, 10) 
        else:
            print("Browser is already running.")

    def login(self, username, password):
        """
        Action: Log in to the SAP system
        """
        try:
            self._start_driver_if_not_running()
            
            # TODO: Replace with your actual Orise SAP login URL
            login_url = 'https://aimprd.advapp.de/sap/bc/ui2/flp'
            self.driver.get(login_url)
            
            # TODO: Replace with your actual "username" input field ID
            # Use self.wait to intelligently wait for the element to appear
            user_field = self.wait.until(
                EC.presence_of_element_located((By.ID, 'USERNAME_FIELD-inner'))
            )
            
            # TODO: Replace with your actual "password" input field ID
            pass_field = self.driver.find_element(By.ID, 'PASSWORD_FIELD-inner')
            
            user_field.send_keys(username)
            pass_field.send_keys(password)
            
            # TODO: Replace with your actual "login" button ID
            self.driver.find_element(By.ID, 'LOGIN_LINK').click()
            
            # After login, wait for an element on the new page to appear, to confirm login success
            # TODO: Replace with the ID of an element on your homepage after login (e.g., logo or "welcome" message)
            self.wait.until(
                EC.presence_of_element_located((By.ID, 'shell-header-logo')) 
            )
            
            print(f"Successfully logged in as {username}.")
            return {"status": "success", "message": "Login successful"}
            
        except Exception as e: # Note: Corrected the original 'except' syntax
            print(f"Login failed: {e}")
            return {"status": "error", "message": str(e)}

print("✅ Cell 1/3: SapRobot Action Layer defined.")

# %%
##################################################################
# Cell 2/3: Intent Layer (The "Translator")
# Run this cell to "define" your Ollama translator in memory
##################################################################
print("Defining Cell 2: Ollama Intent Layer...")

# --- 1. Import Ollama tools ---
import ollama
import json

# --- 2. Define the System Prompt (Ollama's "Instruction Manual") ---
SYSTEM_PROMPT = """
You are an assistant that translates natural language into JSON commands to control an SAP automation robot.
The robot has the following available functions (actions):
1. `login(username, password)`: Log in to the system.
2. `open_worklist()`: Open the worklist page.
3. `quit()`: Quit and close the browser.

Your task is to:
1. Identify the user's intent and match it to one of the actions above.
2. Extract the required parameters from the user's speech.
3. **Only output a structured JSON object**, with no chatting, explanations, or extra text.

The JSON structure must be:
{"action": "function_name", "parameters": {"param1": "value1", ...}}

---
Example 1:
User: "Help me log in, username is HSE-DEMO1, password HSE@demo1"
Assistant: {"action": "login", "parameters": {"username": "HSE-DEMO1", "password": "HSE@demo1"}}

Example 2:
User: "Open the work order list"
Assistant: {"action": "open_worklist", "parameters": {}}

Example 3:
User: "Close the browser"
Assistant: {"action": "quit", "parameters": {}}

Example 4:
User: "What's the weather like today"
Assistant: {"action": "unknown", "parameters": {}}
---
"""

# --- 3. Define the function to communicate with Ollama ---
def get_json_command(user_text):
    """
    Sends the user text to Ollama and retrieves a structured JSON command.
    """
    print(f"Sending to Ollama: {user_text}")
    try:
        response = ollama.chat(
            model='PetrosStav/gemma3-tools:4b', # Ensure your Ollama service is running and the gemma model is downloaded
            messages=[
                {'role': 'system', 'content': SYSTEM_PROMPT},
                {'role': 'user', 'content': user_text}
            ],
            options={'temperature': 0.0} # Low temperature to ensure stable output
        )
        
        response_text = response['message']['content']
        print(f"Ollama response: {response_text}")
        
        # Parse the JSON
        command_data = json.loads(response_text)
        return command_data
        
    except json.JSONDecodeError:
        print("Error: Ollama did not return valid JSON")
        return {"action": "error", "message": "Ollama response is not valid JSON"}
    except Exception as e:
        print(f"Ollama call failed: {e}")
        return {"action": "error", "message": str(e)}

# --- Cell 2 End ---
print("✅ Cell 2/3: Ollama Intent Layer has been defined.")

# %%
##################################################################
# Cell 3/3: Dispatcher Layer (Modified: Bypassing Ollama)
# Run this cell to "start" your assistant
##################################################################
print("Running Cell 3: Dispatcher Layer (Bypass Ollama Mode)...")

import json  # <-- We need to import the json library to parse your input
import time  # <-- The time library was already imported in Cell 1, but importing again just in case

# --- 1. Initialize the Robot ---
# (We must run Cell 1 first so that SapRobot is defined here)
# TODO: Replace with the real path to your local chromedriver.exe
DRIVER_PATH = r"D:/chromedriver/chromedriver.exe"  # (!! Make sure this is your real path !!)
try:
    robot = SapRobot(driver_path=DRIVER_PATH)
except NameError:
    print("="*50)
    print("!! Fatal Error: 'SapRobot' is not defined !!")
    print("!! You must run Cell 1 first !!")
    print("="*50)
    raise  # Stop execution

# --- 2. Start the Main Loop ---
print("====================================================")
print("=== SAP Assistant Started (Mode: Direct JSON Command Input) ===")
print("====================================================")
print('Example: {"action": "open_worklist", "parameters": {}}')
print('To exit: type q')

while True:
    try:
        # 1. (New) Get your JSON command string
        json_text = input("Please enter JSON command > ") 
        
        if json_text.lower() == 'q':
            robot.quit()  # Close the browser before exiting
            print("Program exited.")
            break  # End the while loop
            
        if not json_text:
            continue
            
        # 2. (New) Try to parse your input as JSON
        command_json = None
        try:
            command_json = json.loads(json_text)
        except json.JSONDecodeError:
            print(f"!! Format Error: Your input '{json_text}' is not valid JSON.")
            print('!! Please enter strict JSON format, e.g.: {"action": "quit", "parameters": {}}')
            print("--- Waiting for the next command ---")
            continue  # Skip this loop, wait for the next input

        # 3. (Unchanged) Parse and dispatch the command
        action = command_json.get("action")
        params = command_json.get("parameters", {})  # Get parameters, default to an empty dict
        
        result = None
        
        # This is the core "dispatch" logic (same as before)
        if action == "login":
            result = robot.login(**params) 
            
        elif action == "open_worklist":
            result = robot.open_worklist()
            
        elif action == "quit":
            result = robot.quit()
            
        elif action == "unknown":
            print("Sorry, I could not understand this command.")
            
        else:
            print(f"Error: Received unknown action '{action}'")

        # (Optional) Print the robot's execution result
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
        print("Attempting to close browser and exit...")
        robot.quit()
        break


