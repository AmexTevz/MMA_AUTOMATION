from config import config
import os
from dotenv import load_dotenv
import json
import time
from datetime import datetime
from pywinauto.application import Application
import random
import csv
import click
import win32gui
import win32ui
from ctypes import windll
from PIL import Image
try:
    config.validate()
except (ImportError, ValueError) as e:
    print(f"Configuration error: {e}")
    print("Please create a .env file with your credentials. See .env.example for reference.")
    raise

class MenuManagementAutomation:
    def __init__(self):
        self.app = None
        self.main_window = None
        self.app_path = config.app_path
        self.domain = config.domain
        self.user = config.user
        self.password = config.password
        self._checkbox_cache = {}
        self.data_dir = "data"
        self.screenshot_dir = os.path.join(self.data_dir, "screenshots")
        os.makedirs(self.data_dir, exist_ok=True)
        os.makedirs(self.screenshot_dir, exist_ok=True)

    def login(self):
        try:
            print("\nAttempting to login...")

            login_text = f"{self.domain}\\{self.user}"
            print(f"Attempting to login with: {login_text}")

            self.main_window.Edit0.set_text(login_text)
            self.main_window.Edit2.set_text(self.password.replace("'", "").strip())
            self.main_window.Button4.click()

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

    @staticmethod
    def log_change(property_id, rvc_id, item_id, action, run_id, categories=None):
        """Log changes to CSV file"""
        try:
            data_dir = "data"
            if not os.path.exists(data_dir):
                os.makedirs(data_dir)

            csv_path = os.path.join(data_dir, 'change_history.csv')

            with open(csv_path, 'a', newline='') as file:
                writer = csv.writer(file)
                if file.tell() == 0:
                    writer.writerow(['Run ID', 'Store', 'RVC', 'Item ID and Name', 'Action', 'Categories', 'Timestamp'])
                categories_str = ','.join(categories) if categories else ''
                writer.writerow([
                    run_id,
                    property_id,
                    rvc_id,
                    item_id,
                    action,
                    categories_str,
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ])
            print(f"Change logged to {csv_path}")
        except Exception as e:
            print(f"Error logging change: {e}")

    def process_random_item(self, property_id, rvc_id, run_id):
        try:
            menu_items_container = self.main_window.child_window(
                title="MenuManagement.Data.Models.MenuItem",
                control_type="ListItem",
                found_index=0
            ).parent()

            menu_items = list(menu_items_container.children(control_type="ListItem"))
            if not menu_items:
                print("No menu items found")
                return False

            random.shuffle(menu_items)

            for selected_item in menu_items:
                item_texts = [t.window_text() for t in selected_item.children(control_type="Text") if t.window_text()]
                item_id = item_texts[0] if item_texts else "Unknown"

                print(f"\nChecking item {item_id}...")
                selected_item.click_input()

                if self.check_item_conditions() and self.checkbox_operation("Active", 0):
                    print(f"Successfully deactivated {item_id}")
                    self.log_change(property_id, rvc_id, item_id, "Deactivated", run_id)
                    return True

            print("No suitable items found to deactivate")
            return False

        except Exception as e:
            print(f"Error in process_random_item: {e}")
            return False

    def run(self, property_id=None, rvc_id=None):
        try:
            print("Launching application...")
            os.startfile(self.app_path)
            time.sleep(3)

            self.app = Application(backend="uia").connect(title_re=".*Menu Management QA.*")
            self.main_window = self.app.window(title_re=".*Menu Management QA.*")
            self.app.wait_cpu_usage_lower(threshold=20, timeout=10)

            self.login()
            time.sleep(2)

            if property_id and rvc_id:
                self.search_and_select(property_id, rvc_id)
                time.sleep(2)

            return True

        except Exception as e:
            print(f"Error in run: {e}")
            return False

    def take_screenshot(self, reason="exit"):
        try:
            if self.main_window and self.main_window.exists():
                hwnd = self.main_window.handle

                left, top, right, bot = win32gui.GetWindowRect(hwnd)
                width = right - left
                height = bot - top

                hwndDC = win32gui.GetWindowDC(hwnd)
                mfcDC = win32ui.CreateDCFromHandle(hwndDC)
                saveDC = mfcDC.CreateCompatibleDC()
                saveBitMap = win32ui.CreateBitmap()
                saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
                saveDC.SelectObject(saveBitMap)

                # Take the screenshot by copying the window content
                result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 0)

                bmpinfo = saveBitMap.GetInfo()
                bmpstr = saveBitMap.GetBitmapBits(True)
                im = Image.frombuffer(
                    'RGB',
                    (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                    bmpstr, 'raw', 'BGRX', 0, 1)

                # Save the screenshot
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = os.path.join(self.screenshot_dir, f"screenshot_{reason}_{timestamp}.png")
                im.save(filename)
                print(f"Screenshot saved: {filename}")
                win32gui.DeleteObject(saveBitMap.GetHandle())
                saveDC.DeleteDC()
                mfcDC.DeleteDC()
                win32gui.ReleaseDC(hwnd, hwndDC)

        except Exception as e:
            print(f"Error taking screenshot: {e}")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            if exc_type:
                self.take_screenshot("error")
        finally:
            if self.app:
                try:
                    self.app.kill()
                except Exception as e:
                    print(f"Error closing application: {e}")

    def clear_search(self):
        try:
            search_box = self.main_window.child_window(
                control_type="Edit",
                found_index=0
            )
            if search_box.exists():
                search_box.set_text("")
                time.sleep(1)
                return True
        except Exception as e:
            print(f"Error clearing search: {e}")
            return False

    def add_to_category(self, category_name):
        # To be implemented
        pass

    def remove_from_all_categories(self, property_id, rvc_id, run_id):
        try:
            menu_items_container = self.main_window.child_window(
                title="MenuManagement.Data.Models.MenuItem",
                control_type="ListItem",
                found_index=0
            ).parent()
            selected_items = [item for item in menu_items_container.children(control_type="ListItem")
                              if item.is_selected()]

            if not selected_items:
                print("No item selected")
                return False

            selected_item = selected_items[0]
            item_texts = [t.window_text() for t in selected_item.children(control_type="Text") if t.window_text()]
            item_id = item_texts[0] if item_texts else "Unknown"

            # Find and remove categories
            category_items = self.main_window.descendants(
                title="MenuManagement.Data.Models.Category",
                control_type="DataItem"
            )
            removed_categories = []

            for category in category_items:
                checkboxes = category.descendants(control_type="CheckBox")
                if checkboxes:
                    checkbox = checkboxes[0]
                    if checkbox.get_toggle_state() == 1:
                        text_elements = category.descendants(control_type="Text")
                        if text_elements:
                            category_name = text_elements[0].window_text()
                            checkbox.toggle()
                            removed_categories.append(category_name)
                            print(f"Unchecked category: {category_name}")

            if removed_categories:
                print(f"Removed from categories: {', '.join(removed_categories)}")
                self.log_change(property_id, rvc_id, item_id, "Removed from categories", run_id, removed_categories)
                return True

            print("No checked categories found")
            return False

        except Exception as e:
            print(f"Error removing categories: {e}")
            return False

    def check_item_conditions(self):
        return all([
            self.checkbox_operation("Out-of-stock", 0, check_only=True),
            self.checkbox_operation("Item disabled", 0, check_only=True),
            self.checkbox_operation("Not available on any levels", 0, check_only=True)
        ])


def load_tasks():
    with open('tasks.json', 'r') as file:
        data = json.load(file)
        store = data['store']
        # Filter out disabled tasks and add store to each task
        return [{'store': store, **task} for task in data['tasks']
                if not task.get('disabled', False)]


def get_unique_runs():
    try:
        csv_path = os.path.join("data", 'change_history.csv')
        with open(csv_path, 'r') as file:
            reader = csv.DictReader(file)
            runs = {}
            for row in reader:
                run_id = row['Run ID']
                if run_id not in runs:
                    runs[run_id] = {
                        'run_id': run_id,
                        'timestamp': row['Timestamp']
                    }
            return sorted(runs.values(), key=lambda x: x['timestamp'], reverse=True)
    except FileNotFoundError:
        return []


def load_history(run_id=None):
    changes = []
    try:
        csv_path = os.path.join("data", 'change_history.csv')
        with open(csv_path, 'r') as file:
            reader = csv.DictReader(file)
            changes = [row for row in reader
                       if run_id is None or row['Run ID'] == run_id]
    except FileNotFoundError:
        print("No history file found")
    return changes


def process_task(task, automation):
    store = task['store']
    operations = task['operations']
    mode = task.get('mode', 'execute')
    run_id = task.get('run_id')

    if mode == 'revert':
        print("\nAvailable runs to revert:")
        unique_runs = get_unique_runs()
        if not unique_runs:
            print("No runs found to revert")
            return

        for i, run in enumerate(unique_runs, 1):
            print(f"{i}. Run ID: {run['run_id']} - Time: {run['timestamp']}")

        while True:
            choice = input("\nEnter run number to revert (or 'all' for all runs): ")
            if choice.lower() == 'all':
                run_id = None
                break
            try:
                idx = int(choice) - 1
                if 0 <= idx < len(unique_runs):
                    run_id = unique_runs[idx]['run_id']
                    break
            except ValueError:
                pass
            print("Invalid choice. Please try again.")

    history = load_history(run_id) if mode == 'revert' else []

    if mode == 'execute':
        property_id = store[0]
        rvc_id = store[1] if len(store) > 1 else ""
        automation.run(property_id, rvc_id)

        automation.checkbox_operation('Modifiers', 0)
        automation.checkbox_operation('IsActive', 1)

        for operation in operations:
            search_term = operation['search']
            print(f"Using search term: {search_term}")

            if automation.search_menu_items(search_term):
                menu_items_container = automation.main_window.child_window(
                    title="MenuManagement.Data.Models.MenuItem",
                    control_type="ListItem",
                    found_index=0
                ).parent()

                menu_items = list(menu_items_container.children(control_type="ListItem"))
                if menu_items:
                    if operation['action'] == 'deactivate':
                        if automation.process_random_item(property_id, rvc_id, run_id):
                            time.sleep(0.5)  # Brief delay before screenshot
                            automation.take_screenshot("deactivated")
                    elif operation['action'] == 'remove_categories':
                        random.shuffle(menu_items)
                        found_item_with_categories = False
                        for item in menu_items:
                            item.click_input()
                            time.sleep(1)

                            if automation.check_item_conditions():
                                if automation.remove_from_all_categories(property_id, rvc_id, run_id):
                                    found_item_with_categories = True
                                    time.sleep(0.5)
                                    automation.take_screenshot("removed_categories")
                                    break
                                else:
                                    print("Item has no categories, trying next item...")

                        if not found_item_with_categories:
                            print("No salad items with categories found")

                automation.clear_search()

    else:
        property_id = store[0]
        rvc_id = store[1] if len(store) > 1 else ""

        if not automation.run(property_id, rvc_id):
            print(f"Failed to run automation for store {property_id} {rvc_id}")
            return  # Exit if initial run fails

        store_changes = [c for c in history if c['Store'] == property_id]
        for change in reversed(store_changes):
            automation.search_menu_items(change['Item ID and Name'])
            if change['Action'] == 'Deactivated':
                automation.checkbox_operation('Active', 1)
            elif change['Action'] == 'Removed from categories':
                print(f"Restore categories not implemented yet for: {change['Categories']}")


def get_operation_mode():
    print("\nChoose operation mode:")
    print("1. Execute Changes")
    print("2. Revert Changes")

    while True:
        choice = input("\nEnter your choice (1 or 2): ")
        if choice == '1':
            return 'execute'
        elif choice == '2':
            return 'revert'
        else:
            print("Invalid choice. Please enter 1 or 2.")


def main():
    click.clear()
    click.echo("Menu Management Automation")
    click.echo("=" * 25)

    mode = get_operation_mode()
    tasks = load_tasks()

    run_id = datetime.now().strftime('%H%M%S')
    print(f"Run ID: {run_id}")


    combined_task = {
        'store': tasks[0]['store'],
        'operations': [],
        'mode': mode,
        'run_id': run_id
    }

    for task in tasks:
        combined_task['operations'].extend(task['operations'])

    with MenuManagementAutomation() as automation:
        process_task(combined_task, automation)


if __name__ == "__main__":
    main()
