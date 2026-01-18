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

    # ▼▼▼ Core Navigation (Bilingual + App Translation) ▼▼▼
    def open_app(self, personnel_number, operation_filter, order_id, app_name):
        print(f"Executing workflow: open_app for '{app_name}' (Filter Op: {operation_filter})...")
        try:
            self._start_driver_if_not_running()

            # --- Phase 1: Smart Navigation (Search/Suchen) ---
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
                except: pass

                if not on_personnel_screen:
                    print("Step 0: Go to home...")
                    home_url = 'https://aimprd.advapp.de/sap/bc/ui2/flp#Shell-home'
                    if "Shell-home" not in self.driver.current_url: self.driver.get(home_url)
                    time.sleep(3)
                    print("Step 1: Click tile...")
                    try: self.long_wait.until(EC.element_to_be_clickable((By.ID, "__tile1-title-inner"))).click()
                    except: self.driver.find_element(By.CSS_SELECTOR, "[title='Advanced Manufacturing']").click()
                    time.sleep(5)

            # --- Phase 2: Personnel Number ---
            print("Step 2: Check personnel input...")
            try:
                pn_input = WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable(
                    (By.XPATH, "//input[contains(@id, 'inputPerNr')]")
                ))
                pn_input.clear()
                pn_input.send_keys(personnel_number)
                try: self.driver.find_element(By.XPATH, "//*[text()='Senden']").click()
                except: 
                    try: self.driver.find_element(By.XPATH, "//*[text()='Submit']").click()
                    except: pn_input.send_keys(Keys.ENTER)
                WebDriverWait(self.driver, 15).until(EC.presence_of_element_located(
                    (By.XPATH, "//input[@placeholder='Search' or @placeholder='Suchen']")
                ))
            except: print("No personnel input needed, continuing...")

            # --- Phase 3: Filters (EN: Operation / DE: Vorgang) ---
            print("Step 3: Handle filters (Operation/Vorgang Check)...")
            operation_field_visible = False
            
            try:
                self.driver.find_element(By.XPATH, "//input[contains(@aria-label, 'Operation') or contains(@placeholder, 'Operation') or contains(@aria-label, 'Vorgang') or contains(@placeholder, 'Vorgang')]")
                operation_field_visible = True
            except:
                try:
                    self.driver.find_element(By.XPATH, "//*[text()='Operation' or text()='Vorgang']/following::input[1]")
                    operation_field_visible = True
                except: pass

            if not operation_field_visible:
                print("❌ 'Operation/Vorgang' filter not found -> Adapt Filters...")
                self.driver.find_element(By.XPATH, "//*[contains(text(), 'Adapt Filters') or contains(text(), 'Filter anpassen')]").click()
                time.sleep(2)

                search_box = self.wait.until(EC.element_to_be_clickable((
                    By.XPATH, "//input[@type='search' or @placeholder='Search' or @placeholder='Suchen' or @placeholder='Nach Filtern suchen']"
                )))
                search_box.clear()
                
                is_german = "Suchen" in self.driver.page_source or "Filter anpassen" in self.driver.page_source
                search_term = "Vorgang" if is_german else "Operation"
                
                print(f"Searching for filter: {search_term}")
                search_box.send_keys(search_term) 
                time.sleep(2) 

                print(f"Selecting '{search_term}' checkbox...")
                try:
                    target_row_xpath = f"//tr[.//text()[normalize-space()='{search_term}']][.//div[@role='checkbox']]"
                    checkbox = self.driver.find_element(By.XPATH, target_row_xpath + "//div[@role='checkbox']")
                    checkbox.click()
                    print(f"✅ Clicked '{search_term}' checkbox")
                except:
                    print("⚠️ Precise selection failed, trying loop method...")
                    try:
                        rows = self.driver.find_elements(By.XPATH, "//tr[.//div[@role='checkbox']]")
                        for row in rows:
                            if row.text.strip() == search_term:
                                row.find_element(By.XPATH, ".//div[@role='checkbox']").click()
                                break
                    except: pass

                self.driver.find_element(By.XPATH, "//*[text()='OK']").click()
                time.sleep(3)
            else:
                print("✅ 'Operation/Vorgang' filter present, skip Adapt.")

            # --- Phase 3.3: Tab Input Strategy ---
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
            except:
                try: self.driver.find_element(By.XPATH, "//button[.='Go']").click()
                except: self.driver.find_element(By.XPATH, "//*[text()='Start' or text()='Go']").click()
            time.sleep(3)

            # 5. Click Order Row
            print(f"Step 5: Click Order {order_id}...")
            row = self.long_wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"//*[text()='{order_id}']/ancestor::tr[1]")
            ))
            row.click()
            time.sleep(2)

            # --- 6. Open Final App (🟢 修复点：自动翻译 App 名字) ---
            print(f"Step 6: Open App '{app_name}'...")
            
            # 定义 App 名字的双语映射 (EN -> DE)
            # 参考你的截图 image_44c90d.png
            app_translations = {
                "Capture of production data": "Erfassung von Produktionsdaten",
                "Business Documents": "Business Documents", # 似乎德语版也叫这个
                "Production Monitor": "Produktionsmonitor"
            }
            
            de_name = app_translations.get(app_name, app_name) # 获取德语名，如果没有就用原名
            
            print(f"   -> Looking for EN: '{app_name}' OR DE: '{de_name}'")
            
            # 构建一个同时查找英文名和德文名的 XPath
            # 只要 Tile 的 Title 或 Text 包含其中一个名字即可
            xpath_locator = f"//*[@title='{app_name}' or @title='{de_name}' or contains(text(), '{app_name}') or contains(text(), '{de_name}')]"
            
            try:
                self.long_wait.until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, f"[title='{app_name}'], [title='{de_name}']")
                )).click()
            except:
                try:
                    # 如果 CSS 失败，尝试 XPath 模糊匹配
                    self.driver.find_element(By.XPATH, xpath_locator).click()
                except Exception as e:
                    print(f"❌ Failed to open app tile via Title or Text. Error: {e}")
                    # 最后尝试：只搜前几个词
                    short_name = app_name.split(' ')[0]
                    self.driver.find_element(By.XPATH, f"//*[contains(text(), '{short_name}')]").click()

            print(f"Clicked '{app_name}' / '{de_name}' Tile!")
            return {"status": "success", "message": f"Opened: {app_name}"}

        except Exception as e:
            print(f"!! Workflow failed: {e}")
            return {"status": "error", "message": str(e)}

    

    # ▼▼▼ Qty Confirmation (Bilingual: EN/DE) ▼▼▼
    def qty_confirmation(self, personnel_number="1", operation_filter="", order_id=None,
                         yield_qty="0", scrap_qty="0", rework_qty="0", reason=""):
        """Books quantities (Yield/Scrap/Rework) and optional reason via value help.
        Designed to not affect other use cases (called only when action == 'qty_confirmation')."""
        print(f"Executing workflow: qty_confirmation | Order={order_id} Op={operation_filter} "
              f"Yield={yield_qty} Scrap={scrap_qty} Rework={rework_qty} Reason={reason}")

        try:
            self._start_driver_if_not_running()

            # ---------- helpers ----------
            def _scroll_into_view(el):
                try:
                    self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    time.sleep(0.15)
                except:
                    pass

            def _safe_click(el):
                try:
                    _scroll_into_view(el)
                    el.click()
                    return True
                except:
                    try:
                        self.driver.execute_script("arguments[0].click();", el)
                        return True
                    except:
                        return False

            def press_esc(times=2):
                try:
                    for _ in range(times):
                        ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                        time.sleep(0.2)
                except:
                    pass

            def close_any_popups():
                press_esc(2)
                close_xpaths = [
                    "//button[@title='Close' or @title='Schließen']",
                    "//span[@title='Close' or @title='Schließen']/ancestor::button[1]",
                    "//button[contains(@aria-label,'Close') or contains(@aria-label,'Schließen')]",
                ]
                for xp in close_xpaths:
                    try:
                        for b in self.driver.find_elements(By.XPATH, xp):
                            if b.is_displayed() and b.is_enabled():
                                _safe_click(b)
                                time.sleep(0.2)
                    except:
                        pass

            def click_bilingual_button(texts):
                for t in texts:
                    try:
                        btn = WebDriverWait(self.driver, 6).until(
                            EC.element_to_be_clickable((By.XPATH, f"//button[normalize-space()='{t}']"))
                        )
                        return _safe_click(btn)
                    except:
                        pass
                return False

            def get_input_for_field(field_names):
                # 1) aria-label
                for name in field_names:
                    try:
                        el = WebDriverWait(self.driver, 2).until(
                            EC.presence_of_element_located((By.XPATH, f"//input[contains(@aria-label,'{name}')]"))
                        )
                        if el.is_displayed():
                            return el
                    except:
                        pass

                # 2) by text nearby
                for name in field_names:
                    for txt in [f"{name}:", name]:
                        xps = [
                            f"//*[normalize-space()='{txt}']/following::input[contains(@id,'-inner')][1]",
                            f"//*[normalize-space()='{txt}']/ancestor::*[contains(@class,'sapUiFormElement')][1]//input[contains(@id,'-inner')]",
                        ]
                        for xp in xps:
                            try:
                                el = WebDriverWait(self.driver, 2).until(EC.presence_of_element_located((By.XPATH, xp)))
                                if el.is_displayed():
                                    return el
                            except:
                                pass
                return None

            def set_input_value(inp, val, confirm_tab=True):
                if not inp:
                    return False
                _scroll_into_view(inp)
                try:
                    inp.click()
                except:
                    self.driver.execute_script("arguments[0].click();", inp)
                time.sleep(0.05)
                inp.send_keys(Keys.CONTROL, 'a')
                inp.send_keys(Keys.BACKSPACE)
                inp.send_keys(str(val))
                if confirm_tab:
                    inp.send_keys(Keys.TAB)
                    time.sleep(0.15)
                return True

            def open_reason_value_help():
                reason_inp = get_input_for_field(["Reason", "Grund"])
                if not reason_inp:
                    print("❌ Reason/Grund input not found")
                    return False

                _scroll_into_view(reason_inp)
                time.sleep(0.1)

                # Try F4
                try:
                    reason_inp.click()
                    time.sleep(0.05)
                    reason_inp.send_keys(Keys.F4)
                    WebDriverWait(self.driver, 2).until(
                        EC.presence_of_element_located((By.XPATH, "//*[contains(.,'Select Reason') or contains(.,'Abweichungsgrund')]"))
                    )
                    return True
                except:
                    pass

                # Click value help icon (two squares)
                icon_xps = [
                    "./following::button[1]",
                    "./following::*[contains(@class,'sapMInputBaseIconContainer')][1]",
                    "./following::span[contains(@class,'sapMInputValHelp')][1]",
                    ".//button[.//span[contains(@class,'sapMInputValHelp')]]",
                    ".//*[contains(@class,'sapMInputBaseIcon')]",
                ]
                for xp in icon_xps:
                    try:
                        el = reason_inp.find_element(By.XPATH, xp)
                        if el.is_displayed():
                            return _safe_click(el)
                    except:
                        pass
                return False

            def select_reason(reason_text):
                raw = (reason_text or '').strip()
                if not raw:
                    return False

                # Detect language of popup
                page = ''
                try:
                    page = self.driver.page_source
                except:
                    page = ''
                popup_is_german = ('Abweichungsgrund' in page) or ('auswählen' in page)

                # Normalize + map EN -> DE only if popup is German
                t = raw.lower()
                if popup_is_german:
                    reason_map = {
                        'machine malfunction': 'maschinenschaden',
                        'machine failure': 'maschinenschaden',
                        'operating error': 'bedienungsfehler',
                        'defective material': 'materialfehler',
                        'waste': 'abfall',
                    }
                    t = reason_map.get(t, t)

                # If user already passed German words, accept them
                # (e.g., 'Maschinenschaden')
                t = t.strip().lower()

                # Type into search (no TAB)
                try:
                    search = WebDriverWait(self.driver, 2).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Search' or @placeholder='Suchen']"))
                    )
                    search.click()
                    time.sleep(0.05)
                    search.send_keys(Keys.CONTROL, 'a')
                    search.send_keys(Keys.BACKSPACE)
                    search.send_keys(t)
                    time.sleep(0.2)
                except:
                    pass

                # Click matching row (LI or TR). Use contains on lowercase text.
                li_xpath = (
                    "//li[.//*[contains(translate(normalize-space(),"
                    " 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"
                    f" '{t}')]]"
                )
                tr_xpath = (
                    "//tr[.//*[contains(translate(normalize-space(),"
                    " 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"
                    f" '{t}')]]"
                )

                for xp in [li_xpath, tr_xpath]:
                    try:
                        el = WebDriverWait(self.driver, 6).until(EC.element_to_be_clickable((By.XPATH, xp)))
                        _safe_click(el)
                        time.sleep(0.2)
                        # Some UI5 dialogs need Enter to confirm selection
                        try:
                            ActionChains(self.driver).send_keys(Keys.ENTER).perform()
                        except:
                            pass
                        return True
                    except:
                        pass

                return False

            # ---------- MAIN FLOW ----------
            # Navigate to correct order/op using existing open_app
            nav = self.open_app(
                personnel_number=personnel_number,
                operation_filter=operation_filter,
                order_id=order_id,
                app_name="Capture of production data",
            )
            if isinstance(nav, dict) and nav.get('status') == 'error':
                return nav

            time.sleep(1.2)
            close_any_popups()

            # Open Qty Confirmation / Mengenrückmeldung
            if not click_bilingual_button(["Qty Confirmation", "Mengenrückmeldung"]):
                return {"status": "error", "message": "Could not click Qty Confirmation/Mengenrückmeldung"}

            time.sleep(0.8)
            close_any_popups()

            # Fill fields
            y_inp = get_input_for_field(["Yield", "Gutmenge"])
            s_inp = get_input_for_field(["Scrap", "Ausschuss", "Ausschuß"])
            r_inp = get_input_for_field(["Rework", "Nacharbeit"])

            ok1 = set_input_value(y_inp, yield_qty)
            ok2 = set_input_value(s_inp, scrap_qty)
            ok3 = set_input_value(r_inp, rework_qty)

            if not (ok1 and ok2 and ok3):
                # TAB fallback: start at Yield, then type in active elements
                y = y_inp
                if not y:
                    return {"status": "error", "message": "Could not find Yield/Gutmenge input"}
                actions = ActionChains(self.driver)

                set_input_value(y, yield_qty)  # ends with TAB
                time.sleep(0.15)

                # Now at Scrap
                self.driver.switch_to.active_element.send_keys(Keys.CONTROL, 'a')
                self.driver.switch_to.active_element.send_keys(Keys.BACKSPACE)
                self.driver.switch_to.active_element.send_keys(str(scrap_qty))
                time.sleep(0.15)

                # TAB -> Rework
                actions.send_keys(Keys.TAB).perform()
                time.sleep(0.15)

                self.driver.switch_to.active_element.send_keys(Keys.CONTROL, 'a')
                self.driver.switch_to.active_element.send_keys(Keys.BACKSPACE)
                self.driver.switch_to.active_element.send_keys(str(rework_qty))
                time.sleep(0.15)

            # Reason via value help
            if reason and str(reason).strip():
                close_any_popups()
                if not open_reason_value_help():
                    return {"status": "error", "message": "Could not open Reason/Grund value help"}
                if not select_reason(str(reason).strip()):
                    return {"status": "error", "message": f"Could not select reason: {reason}"}
                time.sleep(0.4)
                close_any_popups()

            # Click Save (EN) / Speichern / Sichern (DE)
            if not click_bilingual_button(["Save", "Speichern", "Sichern"]):
                # fallback xpath with bdi
                try:
                    save_btn = WebDriverWait(self.driver, 6).until(
                        EC.element_to_be_clickable((
                            By.XPATH,
                            "//button[.//bdi[normalize-space()='Save' or normalize-space()='Speichern' or normalize-space()='Sichern'] "
                            "or normalize-space()='Save' or normalize-space()='Speichern' or normalize-space()='Sichern']"
                        ))
                    )
                    if not _safe_click(save_btn):
                        return {"status": "error", "message": "Could not click Save/Speichern/Sichern"}
                except Exception as e:
                    return {"status": "error", "message": f"Could not find Save/Speichern/Sichern button: {e}"}

            return {"status": "success", "message": "Qty confirmation submitted"}

        except Exception as e:
            import traceback
            traceback.print_exc()
            return {"status": "error", "message": f"qty_confirmation crashed: {e}"}

# ▼▼▼ Intelligent Processing (Bilingual: EN/DE) ▼▼▼
    def process_order_step(self, personnel_number, operation_filter, order_id, phase, action_type):
        print(f"Executing: {action_type} {phase} | Order: {order_id} | OpFilter: {operation_filter}")
        self._start_driver_if_not_running()
        
        # Bilingual Mapping
        phase_map = {
            "setup": ["Setup", "Rüsten"],
            "processing": ["Processing", "Bearbeitung", "Bearbeiten", "Verarbeitung"],
            "teardown": ["Teardown", "Abrüsten"]
        }
        
        action_map = {
            "start": ["Start", "starten", "Starten"],
            "end": ["Finish", "End", "beenden", "Beenden", "abschließen"]
        }

        # Build list of possible button texts
        possible_texts = []
        for v in action_map.get(action_type, []):
            for p in phase_map.get(phase, []):
                possible_texts.append(f"{v} {p}") # EN order
                possible_texts.append(f"{p} {v}") # DE order

        print(f"🔍 Looking for buttons: {possible_texts}")

        # ▼▼▼ 🟢 Triple-Click Strategy ▼▼▼
        def try_click_button(xpath_locator):
            try:
                btn = WebDriverWait(self.driver, 1.0).until(EC.element_to_be_clickable((By.XPATH, xpath_locator)))
                print(f"⚡️ Found button '{xpath_locator}', clicking...")
                self.driver.execute_script("arguments[0].style.border='3px solid red'", btn)
                time.sleep(0.5) 
                
                try: ActionChains(self.driver).move_to_element(btn).click().perform(); return True
                except: pass
                try: btn.click(); return True
                except: pass
                self.driver.execute_script("arguments[0].click();", btn); return True
            except: return False

        def clear_popups():
            try:
                ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                close_btns = self.driver.find_elements(By.XPATH, "//button[@title='Close' or @title='Schließen']")
                for btn in close_btns:
                    if btn.is_displayed(): btn.click()
            except: pass

        # --- Smart Check ---
        looks_like_correct_page = False
        if operation_filter:
            try:
                if len(self.driver.find_elements(By.XPATH, f"//*[(contains(text(), 'Operation') or contains(text(), 'Vorgang')) and contains(text(), '{operation_filter}')]")) > 0:
                    print(f"✅ Screen matches Op/Vorgang {operation_filter}.")
                    looks_like_correct_page = True
            except: pass
        
        if not looks_like_correct_page and order_id:
             try:
                if order_id in self.driver.page_source:
                    print(f"✅ Screen matches Order/Auftrag {order_id}.")
                    looks_like_correct_page = True
             except: pass

        if looks_like_correct_page:
            clear_popups()
            for txt in possible_texts:
                if try_click_button(f"//button[.//text()='{txt}']"):
                    return {"status": "success", "message": f"Executed {txt} (Instant)"}
            
            if action_type == "end":
                print(f"⚠️ Attempting other Finish/Beenden buttons...")
                for p_key in ["setup", "processing", "teardown"]:
                    for v in action_map["end"]:
                        for p_val in phase_map[p_key]:
                            txts = [f"{v} {p_val}", f"{p_val} {v}"]
                            for t in txts:
                                if try_click_button(f"//button[.//text()='{t}']"):
                                    return {"status": "success", "message": f"Executed {t} (Auto-Fix)"}

            print("❌ Button not found, preparing to refresh...")

        # --- Fallback Navigation ---
        print("🔄 Executing full navigation flow...")
        # 注意：这里我们依然传英文名，但在 open_app 内部会自动翻译成德文
        nav_result = self.open_app(
            personnel_number, 
            operation_filter, 
            order_id, 
            "Capture of production data" 
        )
        if nav_result["status"] == "error": return nav_result

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
            (By.XPATH, "//span[contains(@class, 'sapUiIcon') and @data-sap-ui-icon-content='']")
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
            (By.XPATH, "//span[contains(@class, 'sapUiIcon') and @data-sap-ui-icon-content='']")
        ]
        return self._click_nav_button(strategies, "Previous/Vorheriger")

    def _click_nav_button(self, strategies, name):
        self._start_driver_if_not_running()
        print(f"🔍 Starting search for {name} button...")
        try: ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
        except: pass

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
                    try: ActionChains(self.driver).move_to_element(target_btn).click().perform(); return {"status": "success", "message": f"Clicked {name}"}
                    except: pass
                    try: target_btn.click(); return {"status": "success", "message": f"Clicked {name}"}
                    except: pass
                    self.driver.execute_script("arguments[0].click();", target_btn)
                    time.sleep(3)
                    return {"status": "success", "message": f"Clicked {name}"}
            except: continue
                
        return {"status": "error", "message": f"Not found {name}"}

    def quit(self):
        if self.driver: self.driver.quit(); self.driver = None

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
            if not output_string: return
            
            s = output_string.find('{')
            e = output_string.rfind('}')
            if s == -1 or e == -1: return
            json_str = output_string[s:e+1]
            print(f"🔍 Extracting: {json_str}") 
            
            try:
                command_json = json.loads(json_str)
            except json.JSONDecodeError:
                print("⚠️ JSON Error, applying auto-fix...")
                try: command_json = json.loads(json_str + "}")
                except: return

    except Exception: return

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
            result = robot.process_order_step(
                personnel_number=params.get("personnel_number", "1"),
                operation_filter=params.get("operation_filter", ""),
                order_id=params.get("order_id"),
                phase=params.get("phase", "processing"),
                action_type=p_type,
            )

        elif action == "qty_confirmation":
            result = robot.qty_confirmation(
                personnel_number=params.get("personnel_number", "1"),
                operation_filter=params.get("operation_filter", ""),
                order_id=params.get("order_id"),
                yield_qty=params.get("yield_qty", "0"),
                scrap_qty=params.get("scrap_qty", "0"),
                rework_qty=params.get("rework_qty", "0"),
                reason=params.get("reason", ""),
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


