try:
    from config import USERNAME, PASSWORD, APP_PATH
except ImportError:
    print("Please create a config.py file with your credentials. See config.example.py for reference.")
    raise
from pywinauto.application import Application
import time
import os
import random


class MenuManagementAutomation:
    def __init__(self):
        self.app = None
        self.main_window = None
        self.appref_path = APP_PATH
        self.username = USERNAME
        self.password = PASSWORD

    def login(self):
        try:
            print("\nAttempting to login...")
            self.main_window.Edit0.set_text(self.username)
            self.main_window.Edit2.set_text(self.password)
            self.main_window.Button4.click()
            time.sleep(3)
            return True
        except Exception as e:
            print(f"Error during login: {e}")
            return False

    def find_checkbox(self, name):
        try:
            checkbox = self.main_window.child_window(title=name, control_type="CheckBox")
            if checkbox.exists():
                print(f"Found checkbox with title '{name}'")
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

    def process_random_item(self):

        try:
            self.checkbox_operation('Modifiers', 0)
            self.checkbox_operation('IsActive', 1)
            menu_items = self.main_window.child_window(
                title="MenuManagement.Data.Models.MenuItem",
                control_type="ListItem",
                found_index=0
            ).parent().children(control_type="ListItem")

            menu_items = list(menu_items)
            checked_items = set()

            while checked_items != set(range(len(menu_items))):
                available_indices = set(range(len(menu_items))) - checked_items
                random_index = random.choice(list(available_indices))
                checked_items.add(random_index)

                selected_item = menu_items[random_index]
                item_texts = [t.window_text() for t in selected_item.children(control_type="Text") if t.window_text()]
                item_id = item_texts[0] if item_texts else "Unknown"

                print(f"\nChecking item {item_id}...")
                selected_item.click_input()
                if (self.checkbox_operation("Out-of-stock",0, check_only=True) and
                        self.checkbox_operation("Item disabled", 0,check_only=True) and
                        self.checkbox_operation("Not available on any levels", 0,check_only=True)):
                    self.checkbox_operation("Active", 0)
                print("deactivated")

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
                        self.process_random_item()
        except Exception as e:
            print(f"Error in run: {e}")


def main():
    automation = MenuManagementAutomation()
    automation.run(33, 810)


if __name__ == "__main__":
    main()

