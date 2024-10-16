import pytest
import allure
from pywinauto.application import Application
import os
import time
from PIL import ImageGrab
import io

def take_screenshot(app_window):
    # Get the window rectangle (position and size)
    rect = app_window.rectangle()

    # Capture screenshot of the specific window region
    screenshot = ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom))
    img_bytes = io.BytesIO()
    screenshot.save(img_bytes, format='PNG')
    return img_bytes.getvalue()

@allure.title("Open Notepad, Write Text, Save, and Close Notepad")
def test_notepad_automation():
    app = None

    try:
        # Step 1: Open Notepad
        with allure.step("Step 1: Open Notepad"):
            app = Application().start("notepad.exe")
            notepad = app.UntitledNotepad
            allure.attach("Notepad Window Title", notepad.window_text())

            expected_result = "Notepad should open with title 'Untitled - Notepad'"
            actual_result = notepad.window_text()
            allure.attach(f"Expected: {expected_result}", name="Expected Result", attachment_type=allure.attachment_type.TEXT)
            allure.attach(f"Actual: {actual_result}", name="Actual Result", attachment_type=allure.attachment_type.TEXT)

            assert actual_result == "Untitled - Notepad", "Notepad did not open correctly"

        # Step 2: Write text in Notepad
        with allure.step("Step 2: Write 'I'm just writing' in Notepad"):
            notepad.Edit.type_keys("I'm just writing", with_spaces=True)
            typed_text = notepad.Edit.window_text()
            allure.attach("Text Written", "I'm just writing")

            expected_result = "I'm just writing"
            actual_result = typed_text
            allure.attach(f"Expected: {expected_result}", name="Expected Result", attachment_type=allure.attachment_type.TEXT)
            allure.attach(f"Actual: {actual_result}", name="Actual Result", attachment_type=allure.attachment_type.TEXT)

            assert expected_result in actual_result, f"Text was not written correctly. Found: {actual_result}"

        # Step 3: Save the file as 'sri.txt'
        with allure.step("Step 3: Save the file with the name 'sri.txt'"):
            notepad.menu_select("File -> Save As")
            app.SaveAs.Edit.type_keys("sri.txt")
            app.SaveAs.Save.click()
            time.sleep(2)  # Wait for file save dialog to close
            allure.attach("Filename", "sri.txt")

            # Specify the full path where the file is saved
            saved_file_path = os.path.join(os.path.expanduser("~"), "Documents", "sri.txt")
            time.sleep(1)
            expected_result = f"File should be saved at {saved_file_path}"
            actual_result = os.path.exists(saved_file_path)
            allure.attach(f"Expected: {expected_result}", name="Expected Result", attachment_type=allure.attachment_type.TEXT)
            allure.attach(f"Actual: {actual_result}", name="Actual Result", attachment_type=allure.attachment_type.TEXT)

            assert actual_result, f"File was not saved at {saved_file_path}"

        # Step 4: Close Notepad
        with allure.step("Step 4: Close Notepad"):
            notepad.menu_select("File -> Exit")
            assert not app.is_process_running(), "Notepad did not close"

    except Exception as e:
        # Attach a screenshot and the exception details in case of failure
        if app:
            allure.attach(take_screenshot(app.UntitledNotepad), name="Error Screenshot", attachment_type=allure.attachment_type.PNG)
        allure.attach(str(e), name="Error Message", attachment_type=allure.attachment_type.TEXT)
        raise  # Re-raise the exception to allow the test to fail normally

    

