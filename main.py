#
# from pywinauto.application import Application
# import time
# import os
#
#
# class MenuManagementAutomation:
#     def __init__(self):
#         self.app = None
#         self.main_window = None
#         self.appref_path = r"C:\Users\atevzadze\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\HMSHost\Digital\Menu Management QA.appref-ms"
#         self.username = r"resource\atevzadze"
#         self.password = "Teklaeli*123456"
#
#     def login(self):
#         try:
#             print("\nAttempting to login...")
#             self.main_window.Edit0.set_text(self.username)
#             self.main_window.Edit2.set_text(self.password)
#             self.main_window.Button4.click()
#             time.sleep(3)
#             return True
#         except Exception as e:
#             print(f"Error during login: {e}")
#             return False
#
#     def search_and_select(self, property_id, rvc_id):
#         try:
#             print("Looking for search field...")
#             search_text = self.main_window.child_window(title="Search Revenue Centers...", control_type="Text")
#             if search_text.exists():
#                 parent = search_text.parent()
#                 edit_controls = parent.children(control_type="Edit")
#                 if edit_controls:
#                     print(f"Typing {property_id} to filter...")
#                     edit_controls[0].set_text(f'{property_id} {rvc_id}')
#                     time.sleep(1)  # Wait for filter results
#
#                     # Look for the specific custom control with column index
#                     print("Looking for target item...")
#                     custom_item = self.main_window.child_window(
#                         title="Item: MenuManagement.Data.Models.RevenueCenter, Column Display Index: 0",
#                         control_type="Custom",
#                         found_index=0
#                     )
#
#                     if custom_item.exists():
#                         print("Found custom item, looking for text...")
#                         text_controls = custom_item.children(control_type="Text")
#                         print(f"Found {len(text_controls)} text controls")
#                         for text in text_controls:
#                             text_content = text.window_text()
#                             print(f"Text content: '{text_content}'")
#
#                             if str(rvc_id) in text_content:
#                                 print("Found target text, double clicking parent...")
#                                 parent = text.parent().parent()  # Go up two levels to get to the DataItem
#                                 parent.double_click_input()
#                                 return True
#                         print("Target text not found")
#                     else:
#                         print("Target item not found")
#                     return False
#         except Exception as e:
#             print(f"Error in search and select: {e}")
#             return False
#
#     def run(self, property_id=142, rvc_id=208):
#         try:
#             print("Launching application...")
#             os.startfile(self.appref_path)
#             time.sleep(2)
#
#             self.app = Application(backend="uia").connect(title_re=".*Menu Management QA.*")
#             self.main_window = self.app.window(title_re=".*Menu Management QA.*")
#
#             if self.login():
#                 print("Login successful")
#
#                 # Click Location Selection button
#                 location_button = self.main_window.child_window(
#                     title="Location Selection",
#                     control_type="Button"
#                 )
#                 location_button.click()
#                 time.sleep(2)
#
#                 self.search_and_select(property_id, rvc_id)
#
#         except Exception as e:
#             print(f"Error in run: {e}")
#
#
# def main():
#     automation = MenuManagementAutomation()
#     # Example usage with different values:
#     automation.run(142, 208)  # You can change these values as needed
#     # Or use default values:
#     # automation.run()
#
#
# if __name__ == "__main__":
#     main()


# from pywinauto.application import Application
# import time
# import os


# class MenuManagementAutomation:
#     def __init__(self):
#         self.app = None
#         self.main_window = None
#         self.appref_path = r"C:\Users\atevzadze\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\HMSHost\Digital\Menu Management QA.appref-ms"
#         self.username = r"resource\atevzadze"
#         self.password = "Teklaeli*123456"
#
#     def login(self):
#         try:
#             print("\nAttempting to login...")
#             self.main_window.Edit0.set_text(self.username)
#             self.main_window.Edit2.set_text(self.password)
#             self.main_window.Button4.click()
#             time.sleep(1)
#             return True
#         except Exception as e:
#             print(f"Error during login: {e}")
#             return False
#
#     def search_and_select(self, property_id, rvc_id):
#         try:
#             print("Looking for search field...")
#             search_box = self.main_window.child_window(
#                 control_type="Edit",
#                 found_index=0
#             )
#
#             if search_box.exists():
#                 print(f"Typing {property_id} {rvc_id} to filter...")
#                 search_box.set_text(f'{property_id} {rvc_id}')
#                 time.sleep(1)  # Wait for filter
#
#                 # Find the dataitem which is two levels up from our custom control
#                 dataitem = self.main_window.child_window(
#                     title="MenuManagement.Data.Models.RevenueCenter",
#                     control_type="DataItem"
#                 )
#                 dataitem.double_click_input()
#                 return True
#
#         except Exception as e:
#             print(f"Error in search and select: {e}")
#             return False
#
#     def run(self, property_id=142, rvc_id=208):
#         try:
#             print("Launching application...")
#             os.startfile(self.appref_path)
#             time.sleep(2)  # Give app time to start
#
#             self.app = Application(backend="uia")
#             self.app.wait_cpu_usage_lower(threshold=5, timeout=10)
#
#             self.app = self.app.connect(title_re=".*Menu Management QA.*")
#             self.main_window = self.app.window(title_re=".*Menu Management QA.*")
#
#             if self.login():
#                 print("Login successful")
#                 time.sleep(3)  # Keep original wait after login
#                 self.search_and_select(property_id, rvc_id)
#
#         except Exception as e:
#             print(f"Error in run: {e}")
#
# def main():
#     automation = MenuManagementAutomation()
#     # Example usage with different values:
#     automation.run(142, 208)  # You can change these values as needed
#     # Or use default values:
#     # automation.run()
#
#
# if __name__ == "__main__":
#     main()










# from pywinauto.application import Application
# import time
# import os
#
#
# class MenuManagementAutomation:
#     def __init__(self):
#         self.app = None
#         self.main_window = None
#         self.appref_path = r"C:\Users\atevzadze\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\HMSHost\Digital\Menu Management QA.appref-ms"
#         self.username = r"resource\atevzadze"
#         self.password = "Teklaeli*123456"
#
#     def login(self):
#         try:
#             print("\nAttempting to login...")
#             self.main_window.Edit0.set_text(self.username)
#             self.main_window.Edit2.set_text(self.password)
#             self.main_window.Button4.click()
#             time.sleep(3)
#             return True
#         except Exception as e:
#             print(f"Error during login: {e}")
#             return False
#
#     def search_and_select(self, property_id, rvc_id):
#         try:
#             print("Looking for search field...")
#             search_box = self.main_window.child_window(
#                 control_type="Edit",
#                 found_index=0
#             )
#
#             if search_box.exists():
#                 print(f"Typing {property_id} {rvc_id} to filter...")
#                 search_box.set_text(f'{property_id} {rvc_id}')
#                 time.sleep(1)  # Wait for filter
#
#                 dataitem = self.main_window.child_window(
#                     title="MenuManagement.Data.Models.RevenueCenter",
#                     control_type="DataItem"
#                 )
#                 dataitem.double_click_input()
#                 return True
#         except Exception as e:
#             print(f"Error in search and select: {e}")
#             return False
#
#     def search_menu_items(self, search_term):
#         try:
#             print(f"Looking for menu items search field...")
#             menu_search = self.main_window.child_window(
#                 control_type="Edit",
#                 found_index=0
#             )
#
#             if menu_search.exists():
#                 print(f"Typing '{search_term}' to search menu items...")
#                 menu_search.set_text(search_term)
#                 time.sleep(1)  # Wait for search results
#                 return True
#         except Exception as e:
#             print(f"Error in menu items search: {e}")
#             return False
#
#     def is_item_available(self, item):
#         try:
#             # Check "Item disabled" checkbox
#             disabled_checkbox = item.child_window(
#                 title="Item disabled",
#                 control_type="CheckBox"
#             )
#             if disabled_checkbox.exists() and disabled_checkbox.get_toggle_state() == 1:  # 1 means checked
#                 return False
#
#             # Check "Not available on any levels" checkbox
#             not_available_checkbox = item.child_window(
#                 title="Not available on any levels",
#                 control_type="CheckBox"
#             )
#             if not_available_checkbox.exists() and not_available_checkbox.get_toggle_state() == 1:
#                 return False
#
#             return True
#         except Exception as e:
#             print(f"Error checking item availability: {e}")
#             return True  # Default to showing item if check fails
#
#     def print_menu_items(self):
#         try:
#             print("\nRetrieving menu items...")
#             # Find all list items that represent menu items
#             menu_items = self.main_window.child_window(
#                 title="MenuManagement.Data.Models.MenuItem",
#                 control_type="ListItem",
#                 found_index=0
#             ).parent().children(control_type="ListItem")
#
#             print("\nAvailable Menu Items:")
#             print("-" * 50)
#             for item in menu_items:
#                 try:
#                     # Only process if item is available
#                     if self.is_item_available(item):
#                         # Get all text elements within the list item
#                         texts = item.children(control_type="Text")
#                         if texts:
#                             item_text = [t.window_text() for t in texts if t.window_text()]
#                             print(" | ".join(item_text))
#                 except Exception as e:
#                     print(f"Error processing item: {e}")
#             print("-" * 50)
#
#         except Exception as e:
#             print(f"Error printing menu items: {e}")
#
#
#     def run(self, property_id=142, rvc_id=208):
#         try:
#             print("Launching application...")
#             os.startfile(self.appref_path)
#             time.sleep(3)
#
#             self.app = Application(backend="uia").connect(title_re=".*Menu Management QA.*")
#             self.main_window = self.app.window(title_re=".*Menu Management QA.*")
#             self.app.wait_cpu_usage_lower(threshold=20, timeout=10)
#
#
#             if self.login():
#                 if self.search_and_select(property_id, rvc_id):
#                     time.sleep(2)  # Wait for menu items to load
#                     if self.search_menu_items("burger"):
#                         time.sleep(1)  # Wait for search results
#                         self.print_menu_items()
#
#         except Exception as e:
#             print(f"Error in run: {e}")
#
#
# def main():
#     automation = MenuManagementAutomation()
#     automation.run(33, 810)
#
#
#
# if __name__ == "__main__":
#     main()


#

# from pywinauto.application import Application
# import time
# import os
# import random
#
#
# class MenuManagementAutomation:
#     def __init__(self):
#         self.app = None
#         self.main_window = None
#         self.appref_path = r"C:\Users\atevzadze\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\HMSHost\Digital\Menu Management QA.appref-ms"
#         self.username = r"resource\atevzadze"
#         self.password = "Teklaeli*123456"
#
#     def login(self):
#         try:
#             print("\nAttempting to login...")
#             self.main_window.Edit0.set_text(self.username)
#             self.main_window.Edit2.set_text(self.password)
#             self.main_window.Button4.click()
#             time.sleep(3)
#             return True
#         except Exception as e:
#             print(f"Error during login: {e}")
#             return False
#
#     def search_and_select(self, property_id, rvc_id):
#         try:
#             print("Looking for search field...")
#             search_box = self.main_window.child_window(
#                 control_type="Edit",
#                 found_index=0
#             )
#
#             if search_box.exists():
#                 print(f"Typing {property_id} {rvc_id} to filter...")
#                 search_box.set_text(f'{property_id} {rvc_id}')
#                 time.sleep(1)  # Wait for filter
#
#                 dataitem = self.main_window.child_window(
#                     title="MenuManagement.Data.Models.RevenueCenter",
#                     control_type="DataItem"
#                 )
#                 dataitem.double_click_input()
#                 return True
#         except Exception as e:
#             print(f"Error in search and select: {e}")
#             return False
#
#     def search_menu_items(self, search_term):
#         try:
#             print(f"Looking for menu items search field...")
#             menu_search = self.main_window.child_window(
#                 control_type="Edit",
#                 found_index=0
#             )
#
#             if menu_search.exists():
#                 print(f"Typing '{search_term}' to search menu items...")
#                 menu_search.set_text(search_term)
#                 time.sleep(1)  # Wait for search results
#                 return True
#         except Exception as e:
#             print(f"Error in menu items search: {e}")
#             return False
#
#     def check_item_status(self, item):
#         """Check if item is disabled, not available, or active"""
#         try:
#             print(f"\nChecking item status...")
#             item.click_input()
#             time.sleep(0.5)  # Wait for UI update
#
#             # First look for "Item disabled" checkbox
#             disabled_checkbox = self.main_window.child_window(
#                 title="Item disabled",
#                 control_type="CheckBox",
#                 found_index=0  # Get the first match
#             )
#             if disabled_checkbox.exists():
#                 disabled_state = disabled_checkbox.get_toggle_state()
#                 print(f"Item disabled checkbox state: {disabled_state}")
#                 if disabled_state == 1:  # If checked
#                     return "disabled"
#
#             # Then look for "Not available on any levels" checkbox
#             not_available_checkbox = self.main_window.child_window(
#                 title="Not available on any levels",
#                 control_type="CheckBox",
#                 found_index=0  # Get the first match
#             )
#             if not_available_checkbox.exists():
#                 not_available_state = not_available_checkbox.get_toggle_state()
#                 print(f"Not available checkbox state: {not_available_state}")
#                 if not_available_state == 1:  # If checked
#                     return "not_available"
#
#             # Finally check for Active checkbox
#             active_text = self.main_window.child_window(
#                 title="Active:",
#                 control_type="Text"
#             )
#             if active_text.exists():
#                 active_checkbox = active_text.parent().child_window(
#                     control_type="CheckBox"
#                 )
#                 if active_checkbox.exists():
#                     active_state = active_checkbox.get_toggle_state()
#                     print(f"Active checkbox state: {active_state}")
#                     if active_state == 1:
#                         return "active"
#                     else:
#                         return "inactive"
#
#             return "unknown"
#         except Exception as e:
#             print(f"Error checking item status: {e}")
#             return "unknown"
#
#     def toggle_active_state(self, item):
#         """Toggle the active state of an item"""
#         try:
#             # Find all checkboxes
#             checkboxes = self.main_window.descendants(control_type="CheckBox")
#
#             for checkbox in checkboxes:
#                 parent = checkbox.parent()
#                 if parent:
#                     texts = parent.descendants(control_type="Text")
#                     if any(t.window_text() == "Active:" for t in texts):
#                         initial_state = checkbox.get_toggle_state()
#                         checkbox.click_input()
#                         time.sleep(0.5)
#                         new_state = checkbox.get_toggle_state()
#                         return initial_state != new_state  # True if toggle was successful
#             return False
#         except Exception as e:
#             print(f"Error toggling active state: {e}")
#             return False
#
#     def process_random_item(self):
#         try:
#             # Get all menu items
#             menu_items = self.main_window.child_window(
#                 title="MenuManagement.Data.Models.MenuItem",
#                 control_type="ListItem",
#                 found_index=0
#             ).parent().children(control_type="ListItem")
#
#             menu_items = list(menu_items)
#             checked_items = set()
#
#             while checked_items != set(range(len(menu_items))):
#                 # Randomly select an unchecked item
#                 available_indices = set(range(len(menu_items))) - checked_items
#                 random_index = random.choice(list(available_indices))
#                 checked_items.add(random_index)
#
#                 selected_item = menu_items[random_index]
#                 item_texts = [t.window_text() for t in selected_item.children(control_type="Text") if t.window_text()]
#                 item_id = item_texts[0] if item_texts else "Unknown"
#
#                 print(f"\nChecking item {item_id}...")
#                 status = self.check_item_status(selected_item)
#
#                 if status == "disabled":
#                     print(f"Item {item_id} is disabled, trying another...")
#                     continue
#                 elif status == "not_available":
#                     print(f"Item {item_id} is not available, trying another...")
#                     continue
#                 elif status == "active":
#                     print(f"Item {item_id} is active, attempting to deactivate...")
#                     if self.toggle_active_state(selected_item):
#                         print(f"Successfully deactivated item {item_id}")
#                         return True
#                     else:
#                         print(f"Failed to deactivate item {item_id}, trying another...")
#                 elif status == "inactive":
#                     print(f"Item {item_id} is already inactive, trying another...")
#                 else:
#                     print(f"Item {item_id} status unknown, trying another...")
#
#             print("\nChecked all items but found no active items to deactivate")
#             return False
#
#         except Exception as e:
#             print(f"Error in process_random_item: {e}")
#             return False
#
#     def run(self, property_id=142, rvc_id=208):
#         try:
#             print("Launching application...")
#             os.startfile(self.appref_path)
#             time.sleep(3)
#
#             self.app = Application(backend="uia").connect(title_re=".*Menu Management QA.*")
#             self.main_window = self.app.window(title_re=".*Menu Management QA.*")
#             self.app.wait_cpu_usage_lower(threshold=20, timeout=10)
#
#             if self.login():
#                 if self.search_and_select(property_id, rvc_id):
#                     time.sleep(2)  # Wait for menu items to load
#                     if self.search_menu_items("burger"):
#                         time.sleep(1)  # Wait for search results
#                         self.process_random_item()
#
#         except Exception as e:
#             print(f"Error in run: {e}")
#
#
# def main():
#     automation = MenuManagementAutomation()
#     automation.run(33, 810)
#
#
# if __name__ == "__main__":
#     main()

# from pywinauto.application import Application
# import time
# import os
# import random
#
#
# class MenuManagementAutomation:
#     def __init__(self):
#         self.app = None
#         self.main_window = None
#         self.appref_path = r"C:\Users\atevzadze\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\HMSHost\Digital\Menu Management QA.appref-ms"
#         self.username = r"resource\atevzadze"
#         self.password = "Teklaeli*123456"
#
#     def login(self):
#         try:
#             print("\nAttempting to login...")
#             self.main_window.Edit0.set_text(self.username)
#             self.main_window.Edit2.set_text(self.password)
#             self.main_window.Button4.click()
#             time.sleep(3)
#             return True
#         except Exception as e:
#             print(f"Error during login: {e}")
#             return False
#
#     def search_and_select(self, property_id, rvc_id):
#         try:
#             print("Looking for search field...")
#             search_box = self.main_window.child_window(
#                 control_type="Edit",
#                 found_index=0
#             )
#
#             if search_box.exists():
#                 print(f"Typing {property_id} {rvc_id} to filter...")
#                 search_box.set_text(f'{property_id} {rvc_id}')
#                 time.sleep(1)  # Wait for filter
#
#                 dataitem = self.main_window.child_window(
#                     title="MenuManagement.Data.Models.RevenueCenter",
#                     control_type="DataItem"
#                 )
#                 dataitem.double_click_input()
#                 return True
#         except Exception as e:
#             print(f"Error in search and select: {e}")
#             return False
#
#     def search_menu_items(self, search_term):
#         try:
#             print(f"Looking for menu items search field...")
#             menu_search = self.main_window.child_window(
#                 control_type="Edit",
#                 found_index=0
#             )
#
#             if menu_search.exists():
#                 print(f"Typing '{search_term}' to search menu items...")
#                 menu_search.set_text(search_term)
#                 time.sleep(1)  # Wait for search results
#                 return True
#         except Exception as e:
#             print(f"Error in menu items search: {e}")
#             return False
#
#     def check_item_status(self, item):
#         """Check if item is disabled, not available, or active"""
#         try:
#             print(f"\nChecking item status...")
#             item.click_input()
#             time.sleep(0.5)  # Wait for UI update
#
#             # Track which checkboxes we've already checked using their rectangle coordinates
#             checked_locations = set()
#
#             # Find the disabled checkbox
#             disabled_checkbox = None
#             not_available_checkbox = None
#             active_checkbox = None
#
#             # Find all relevant elements
#             checkboxes = self.main_window.descendants(control_type="CheckBox")
#             text_elements = {elem.window_text(): elem for elem in self.main_window.descendants(control_type="Text")}
#
#             # Find active text location if it exists
#             active_text = text_elements.get("Active:")
#
#             for checkbox in checkboxes:
#                 try:
#                     rect = checkbox.rectangle()
#                     location = (rect.left, rect.top)  # Use position as unique identifier
#
#                     # Skip if we've already checked this location
#                     if location in checked_locations:
#                         continue
#                     checked_locations.add(location)
#
#                     # Get nearest text elements
#                     nearby_texts = [t for t in checkbox.parent().descendants(control_type="Text")
#                                     if abs(t.rectangle().top - rect.top) < 30]  # Within 30 pixels vertically
#
#                     nearby_labels = [t.window_text() for t in nearby_texts]
#
#                     if "Item disabled" in nearby_labels:
#                         disabled_checkbox = checkbox
#                     elif "Not available on any levels" in nearby_labels:
#                         not_available_checkbox = checkbox
#                     elif active_text and abs(rect.top - active_text.rectangle().top) < 30:
#                         active_checkbox = checkbox
#
#                 except Exception as e:
#                     print(f"Error checking checkbox: {e}")
#                     continue
#
#             # Now check states in order
#             if disabled_checkbox and disabled_checkbox.get_toggle_state() == 1:
#                 print("Item is disabled")
#                 return "disabled"
#
#             if not_available_checkbox and not_available_checkbox.get_toggle_state() == 1:
#                 print("Item is not available")
#                 return "not_available"
#
#             if active_checkbox:
#                 is_active = active_checkbox.get_toggle_state() == 1
#                 print(f"Item is {'active' if is_active else 'inactive'}")
#                 return "active" if is_active else "inactive"
#
#             print("Status unknown")
#             return "unknown"
#
#         except Exception as e:
#             print(f"Error in item status check: {e}")
#             return "unknown"
#
#     def toggle_active_state(self, item):
#         """Toggle the active state of an item"""
#         try:
#             # Find all checkboxes
#             checkboxes = self.main_window.descendants(control_type="CheckBox")
#
#             for checkbox in checkboxes:
#                 parent = checkbox.parent()
#                 if parent:
#                     texts = parent.descendants(control_type="Text")
#                     if any(t.window_text() == "Active:" for t in texts):
#                         initial_state = checkbox.get_toggle_state()
#                         checkbox.click_input()
#                         time.sleep(0.5)
#                         new_state = checkbox.get_toggle_state()
#                         return initial_state != new_state  # True if toggle was successful
#             return False
#         except Exception as e:
#             print(f"Error toggling active state: {e}")
#             return False
#
#     def process_random_item(self):
#         try:
#             # Get all menu items
#             menu_items = self.main_window.child_window(
#                 title="MenuManagement.Data.Models.MenuItem",
#                 control_type="ListItem",
#                 found_index=0
#             ).parent().children(control_type="ListItem")
#
#             menu_items = list(menu_items)
#             checked_items = set()
#
#             while checked_items != set(range(len(menu_items))):
#                 # Randomly select an unchecked item
#                 available_indices = set(range(len(menu_items))) - checked_items
#                 random_index = random.choice(list(available_indices))
#                 checked_items.add(random_index)
#
#                 selected_item = menu_items[random_index]
#                 item_texts = [t.window_text() for t in selected_item.children(control_type="Text") if t.window_text()]
#                 item_id = item_texts[0] if item_texts else "Unknown"
#
#                 print(f"\nChecking item {item_id}...")
#                 status = self.check_item_status(selected_item)
#
#                 if status == "disabled":
#                     print(f"Item {item_id} is disabled, trying another...")
#                     continue
#                 elif status == "not_available":
#                     print(f"Item {item_id} is not available, trying another...")
#                     continue
#                 elif status == "active":
#                     print(f"Item {item_id} is active, attempting to deactivate...")
#                     if self.toggle_active_state(selected_item):
#                         print(f"Successfully deactivated item {item_id}")
#                         return True
#                     else:
#                         print(f"Failed to deactivate item {item_id}, trying another...")
#                 elif status == "inactive":
#                     print(f"Item {item_id} is already inactive, trying another...")
#                 else:
#                     print(f"Item {item_id} status unknown, trying another...")
#
#             print("\nChecked all items but found no active items to deactivate")
#             return False
#
#         except Exception as e:
#             print(f"Error in process_random_item: {e}")
#             return False
#
#     def run(self, property_id=142, rvc_id=208):
#         try:
#             print("Launching application...")
#             os.startfile(self.appref_path)
#             time.sleep(3)
#
#             self.app = Application(backend="uia").connect(title_re=".*Menu Management QA.*")
#             self.main_window = self.app.window(title_re=".*Menu Management QA.*")
#             self.app.wait_cpu_usage_lower(threshold=20, timeout=10)
#
#             if self.login():
#                 if self.search_and_select(property_id, rvc_id):
#                     time.sleep(2)  # Wait for menu items to load
#                     if self.search_menu_items("burger"):
#                         time.sleep(1)  # Wait for search results
#                         self.process_random_item()
#
#         except Exception as e:
#             print(f"Error in run: {e}")
#
#
# def main():
#     automation = MenuManagementAutomation()
#     automation.run(33, 810)
#
#
# if __name__ == "__main__":
#     main()
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

