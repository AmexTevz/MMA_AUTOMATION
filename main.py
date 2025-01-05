try:
    from credentials import USERNAME, PASSWORD, APP_PATH
except ImportError:
    print("Please create a config.py file with your credentials. See config.example.py for reference.")
    raise
from pywinauto.application import Application
import time
import os
import random
import csv
from datetime import datetime
from PIL import ImageGrab
import win32gui

WAIT_TIME_LOGIN = 3
WAIT_TIME_LOAD = 2
SEARCH_FILTER_WAIT = 1
CPU_USAGE_THRESHOLD = 20
CPU_USAGE_TIMEOUT = 10

class MenuManagementAutomation:
    def __init__(self):
        self.app = None
        self.main_window = None
        self.appref_path = APP_PATH
        self.username = USERNAME
        self.password = PASSWORD
        self.log_file = f"menu_changes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        self._checkbox_cache = {}
        self.screenshot_dir = "screenshots"
        if not os.path.exists(self.screenshot_dir):
            os.makedirs(self.screenshot_dir)


    def login(self):
        try:
            print("\nAttempting to login...")
            self.main_window.Edit0.set_text(self.username)
            self.main_window.Edit2.set_text(self.password)
            self.main_window.Button4.click()
            time.sleep(WAIT_TIME_LOGIN)
            if self.main_window.exists():
                return True
            return False
        except Exception as e:
            print(f"Error during login: {e}")
            return False

    def find_checkbox(self, name):
        if name in self._checkbox_cache:
            if self._checkbox_cache[name].exists():
                return self._checkbox_cache[name]
            else:
                del self._checkbox_cache[name]
                
        try:
            checkbox = self.main_window.child_window(title=name, control_type="CheckBox")
            if checkbox.exists():
                self._checkbox_cache[name] = checkbox
                return checkbox

            text_element = self.main_window.child_window(
                title=f"{name}:",
                control_type="Text"
            )
            if text_element.exists():
                print(f"Found text element '{name}:', looking for associated checkbox")
                text_rect = text_element.rectangle()
                checkboxes = self.main_window.descendants(control_type="CheckBox")

                for checkbox in checkboxes:
                    checkbox_rect = checkbox.rectangle()
                    if (abs(checkbox_rect.top - text_rect.top) < 10 and
                            checkbox_rect.left > text_rect.left and
                            checkbox_rect.left - text_rect.right < 50):
                        print(f"Found checkbox associated with '{name}'")
                        return checkbox

            print(f"Could not find checkbox for '{name}'")
            return None
        except Exception as e:
            print(f"Error finding checkbox for '{name}': {e}")
            return None

    def checkbox_operation(self, name, target_state=None, check_only=False):
        checkbox = self.find_checkbox(name)
        if checkbox:
            try:
                check_state = checkbox.get_toggle_state()

                if check_only:
                    return check_state == target_state

                print(f"Checkbox '{name}' current state: {check_state}")

                if target_state is not None:
                    if check_state != target_state:
                        print(f"Toggling checkbox '{name}' to state {target_state}")
                        checkbox.toggle()
                        new_state = checkbox.get_toggle_state()
                        print(f"New state: {new_state}")
                    else:
                        print(f"Checkbox '{name}' already in desired state {target_state}")
                    return True

            except Exception as e:
                print(f"Error operating checkbox '{name}': {e}")
        return False

    def search_and_select(self, property_id, rvc_id):
        try:
            print("Looking for search field...")
            search_box = self.main_window.child_window(
                control_type="Edit",
                found_index=0
            )

            if search_box.exists():
                print(f"Typing {property_id} {rvc_id} to filter...")
                search_box.set_text(f'{property_id} {rvc_id}')
                time.sleep(1)  # Wait for filter

                dataitem = self.main_window.child_window(
                    title="MenuManagement.Data.Models.RevenueCenter",
                    control_type="DataItem"
                )
                dataitem.double_click_input()
                return True
        except Exception as e:
            print(f"Error in search and select: {e}")
            return False

    def search_menu_items(self, search_term):
        try:
            print(f"Looking for menu items search field...")
            menu_search = self.main_window.child_window(
                control_type="Edit",
                found_index=0
            )
            if menu_search.exists():
                print(f"Typing '{search_term}' to search menu items...")
                menu_search.set_text(search_term)
                time.sleep(1)
                return True
        except Exception as e:
            print(f"Error in menu items search: {e}")
            return False

    def log_change(self, property_id, rvc_id, item_id, action):
        """Log changes to CSV file"""
        try:
            with open(self.log_file, 'a', newline='') as file:
                writer = csv.writer(file)
                if file.tell() == 0:
                    writer.writerow(['Store', 'RVC', 'Item ID and Name', 'Action', 'Timestamp'])
                writer.writerow([
                    property_id,
                    rvc_id,
                    item_id,
                    action,
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ])
            print(f"Change logged to {self.log_file}")
        except Exception as e:
            print(f"Error logging change: {e}")

    def process_random_item(self, property_id, rvc_id):
        try:
            self.checkbox_operation('Modifiers', 0)
            self.checkbox_operation('IsActive', 1)
            
            # Get menu items once and cache them
            menu_items_container = self.main_window.child_window(
                title="MenuManagement.Data.Models.MenuItem",
                control_type="ListItem",
                found_index=0
            ).parent()
            
            menu_items = list(menu_items_container.children(control_type="ListItem"))
            if not menu_items:
                print("No menu items found")
                return False
                
            # Shuffle the list instead of using random selection
            random.shuffle(menu_items)
            
            for selected_item in menu_items:
                item_texts = [t.window_text() for t in selected_item.children(control_type="Text") if t.window_text()]
                item_id = item_texts[0] if item_texts else "Unknown"

                print(f"\nChecking item {item_id}...")
                selected_item.click_input()
                
                # Combine all checkbox checks into one condition
                all_conditions_met = all([
                    self.checkbox_operation("Out-of-stock", 0, check_only=True),
                    self.checkbox_operation("Item disabled", 0, check_only=True),
                    self.checkbox_operation("Not available on any levels", 0, check_only=True)
                ])
                
                if all_conditions_met and self.checkbox_operation("Active", 0):
                    print(f"Successfully deactivated {item_id}")
                    self.log_change(property_id, rvc_id, item_id, "Deactivated")
                    return True

            print("No suitable items found to deactivate")
            return False

        except Exception as e:
            print(f"Error in process_random_item: {e}")
            return False

    def run(self, property_id=None, rvc_id=None):
        try:
            print("Launching application...")
            os.startfile(self.appref_path)
            time.sleep(3)
            self.app = Application(backend="uia").connect(title_re=".*Menu Management QA.*")
            self.main_window = self.app.window(title_re=".*Menu Management QA.*")
            self.app.wait_cpu_usage_lower(threshold=20, timeout=10)

            if self.login():
                if self.search_and_select(property_id, rvc_id):
                    time.sleep(2)  # Wait for menu items to load
                    if self.search_menu_items("burger"):
                        if self.process_random_item(property_id, rvc_id):
                            print("Successfully processed one item")
                        else:
                            print("No items were processed")
        except Exception as e:
            print(f"Error in run: {e}")

    def take_screenshot(self, reason="exit"):
        """Take a screenshot of the current window"""
        try:
            if self.main_window and self.main_window.exists():
                # Get the window handle
                hwnd = self.main_window.handle
                
                # Get window coordinates
                window_rect = win32gui.GetWindowRect(hwnd)
                
                # Capture the specific region
                screenshot = ImageGrab.grab(window_rect)
                
                # Save the screenshot
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = os.path.join(self.screenshot_dir, f"screenshot_{reason}_{timestamp}.png")
                screenshot.save(filename)
                print(f"Screenshot saved: {filename}")
        except Exception as e:
            print(f"Error taking screenshot: {e}")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            if exc_type:  # If there was an error
                self.take_screenshot("error")
            else:
                self.take_screenshot("normal_exit")
        finally:
            if self.app:
                try:
                    self.app.kill()
                except Exception as e:
                    print(f"Error closing application: {e}")


def main():
    with MenuManagementAutomation() as automation:
        automation.run(33, 810)


if __name__ == "__main__":
    main()
