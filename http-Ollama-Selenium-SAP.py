##################################################################
# SAP Robot – FINAL WORKING VERSION WITH FUZZY MATCHING
##################################################################

print("Loading SapRobot...")

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

from flask import Flask, request, jsonify
import json
import threading
from difflib import SequenceMatcher


##################################################################
# Helper: Normalize text for fuzzy matching
##################################################################
def normalize_text(t):
    """Entfernt Leerzeichen, Sonderzeichen, Trenner und macht lowercase."""
    return "".join(c for c in t.lower() if c.isalnum())


##################################################################
# SapRobot Class
##################################################################
class SapRobot:
    def __init__(self):
        self.driver = None
        self.wait = None
        print("SapRobot initialized.")

    def _start_driver_if_not_running(self):
        if self.driver is None:
            print("Starting Chrome...")

            chrome_options = Options()
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.wait = WebDriverWait(self.driver, 10)

        else:
            print("Chrome is already running.")

    ##################################################################
    # LOGIN
    ##################################################################
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

            self.wait.until(
                EC.presence_of_element_located((By.ID, "shell-header-logo"))
            )

            print("Login success.")
            return {"status": "success", "message": "Login successful"}

        except Exception as e:
            print("Login error:", e)
            return {"status": "error", "message": str(e)}

    ##################################################################
    # QUIT
    ##################################################################
    def quit(self):
        if self.driver is not None:
            try:
                self.driver.quit()
            except:
                pass

            self.driver = None
            self.wait = None

            print("Browser closed.")
            return {"status": "success", "message": "Browser closed"}

        return {"status": "success", "message": "Browser was not running."}

    ##################################################################
    # OPEN APP WITH FUZZY MATCHING
    ##################################################################
    def open_app(self, app_name: str):
        try:
            self._start_driver_if_not_running()

            print(f"Searching tiles for '{app_name}'")

            cleaned_query = normalize_text(app_name)
            print("Normalized query:", cleaned_query)

            tiles = self.wait.until(
                EC.presence_of_all_elements_located(
                    (
                        By.CSS_SELECTOR,
                        "[role='button'], .sapUshellTile, .sapUshellTileBase",
                    )
                )
            )

            print(f"Found {len(tiles)} tiles.")

            best_match = None
            best_score = 0.0

            for tile in tiles:
                tile_text_raw = tile.text.strip()
                cleaned_tile = normalize_text(tile_text_raw)

                score = SequenceMatcher(
                    None, cleaned_query, cleaned_tile
                ).ratio()

                print(f"Tile: '{tile_text_raw}' → score {score:.3f}")

                if score > best_score:
                    best_score = score
                    best_match = tile

            if best_score < 0.45:
                print("No good match found.")
                return {"status": "error", "message": f"No tile found for '{app_name}'"}

            print(f"BEST MATCH SCORE: {best_score:.3f} – clicking tile…")
            best_match.click()

            return {"status": "success", "message": f"Opened '{app_name}'"}

        except Exception as e:
            print("open_app error:", e)
            return {"status": "error", "message": str(e)}


##################################################################
# Flask Server for n8n
##################################################################

app = Flask(__name__)
robot = SapRobot()

@app.route("/execute", methods=["POST"])
def execute_command():
    try:
        raw = request.data.decode("utf-8")
        print("JSON received:", raw)

        command = json.loads(raw)

        def run():
            action = command.get("action")
            params = command.get("params", {})

            print("Dispatching:", action, params)

            if action == "login":
                robot.login(**params)
            elif action == "open_app":
                robot.open_app(**params)
            elif action == "quit":
                robot.quit()
            else:
                print("Unknown action:", action)

        threading.Thread(target=run).start()

        return jsonify({"status": "received"})

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


##################################################################
# START SERVER (IMPORTANT!)
##################################################################
if __name__ == "__main__":   # beide Unterstriche!!
    print("Robot Server running at http://localhost:5000/execute")
    app.run(host="localhost", port=5000, debug=False)

