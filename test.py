import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

driver = webdriver.Chrome()

password_change_url = "http://localhost:8080"

excel_file = "testSheet.xlsx"
sheet_name = "Sheet1"

wb = openpyxl.load_workbook(excel_file)
sheet = wb[sheet_name]

column_letter = "C"
username = "jasnagra"

start_row = 4
end_row = 7

try:
    for row_num in range(start_row, end_row + 1):
        print(f"Processing row {row_num}")
        cell_value = sheet[f"{column_letter}{row_num}"].value
        print(cell_value)

        if cell_value is not None:
            password = cell_value

            driver.get(password_change_url)

            time.sleep(2)

            username_field = driver.find_element(By.ID, "user")
            password_field = driver.find_element(By.ID, "password")
            login_button = driver.find_element(By.NAME, "btnSubmit")

            username_field.send_keys(username)
            password_field.send_keys(password)

            login_button.click()

            time.sleep(3)

            second_page_username_field = driver.find_element(
                By.ID,
                "ctl00_mainContent_ChangePassword1_ChangePasswordContainerID_UserName",
            )
            current_password_field = driver.find_element(
                By.ID,
                "ctl00_mainContent_ChangePassword1_ChangePasswordContainerID_CurrentPassword",
            )
            new_password_field = driver.find_element(
                By.ID,
                "ctl00_mainContent_ChangePassword1_ChangePasswordContainerID_NewPassword",
            )
            confirm_new_password_field = driver.find_element(
                By.ID,
                "ctl00_mainContent_ChangePassword1_ChangePasswordContainerID_ConfirmNewPassword",
            )

            next_row_num = row_num + 1
            new_password = sheet[f"{column_letter}{next_row_num}"].value

            second_page_username_field.send_keys(username)
            current_password_field.send_keys(password)
            new_password_field.send_keys(new_password)
            confirm_new_password_field.send_keys(new_password)

            change_password_button = driver.find_element(
                By.ID,
                "ctl00_mainContent_ChangePassword1_ChangePasswordContainerID_ChangePasswordPushButton",
            )
            change_password_button.click()

            time.sleep(3)

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
