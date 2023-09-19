import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

driver = webdriver.Chrome()

password_change_url = "http://localhost:8080"

excel_file = "IDIR_usernames.xlsx"
sheet_name = "Sheet1"

wb = openpyxl.load_workbook(excel_file)
sheet = wb[sheet_name]

column_letter = "A"
password = "Wildfire.2023"
new_password = "Blizzard.2023"

start_row = 2
end_row = 22

try:
    for row_num in range(start_row, end_row + 1):
        print(f"Processing row {row_num}")
        cell_value = sheet[f"{column_letter}{row_num}"].value
        print(cell_value)

        if cell_value is not None:
            username = cell_value

            driver.get(password_change_url)

            wait = WebDriverWait(driver, 6)

            username_field = wait.until(EC.element_to_be_clickable((By.ID, "user")))
            password_field = wait.until(EC.element_to_be_clickable((By.ID, "password")))
            login_button = wait.until(
                EC.element_to_be_clickable((By.NAME, "btnSubmit"))
            )

            username_field.send_keys(username)
            password_field.send_keys(password)

            login_button.click()

            second_page_username_field = wait.until(
                EC.presence_of_element_located(
                    (
                        By.ID,
                        "ctl00_mainContent_ChangePassword1_ChangePasswordContainerID_UserName",
                    )
                )
            )

            current_password_field = wait.until(
                EC.presence_of_element_located(
                    (
                        By.ID,
                        "ctl00_mainContent_ChangePassword1_ChangePasswordContainerID_CurrentPassword",
                    )
                )
            )
            new_password_field = wait.until(
                EC.presence_of_element_located(
                    (
                        By.ID,
                        "ctl00_mainContent_ChangePassword1_ChangePasswordContainerID_NewPassword",
                    )
                )
            )
            confirm_new_password_field = wait.until(
                EC.presence_of_element_located(
                    (
                        By.ID,
                        "ctl00_mainContent_ChangePassword1_ChangePasswordContainerID_ConfirmNewPassword",
                    )
                )
            )

            second_page_username_field.send_keys(username)
            current_password_field.send_keys(password)
            new_password_field.send_keys(new_password)
            confirm_new_password_field.send_keys(new_password)

            change_password_button = wait.until(
                EC.presence_of_element_located(
                    (
                        By.ID,
                        "ctl00_mainContent_ChangePassword1_ChangePasswordContainerID_ChangePasswordPushButton",
                    )
                )
            )
            change_password_button.click()

            driver.delete_all_cookies()
            print(f"Finished processing {row_num}")
        else:
            print(f"Skipping row {row_num} because cell is empty")
except Exception as e:
    print("an unexpected error occured", str(e))
finally:
    try:
        driver.quit()
    except Exception as e:
        print("An error occured while quitting the driver:", str(e))
