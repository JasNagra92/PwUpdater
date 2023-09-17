import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

driver = webdriver.Chrome()

password_change_url = "https://www.pwchange.gov.bc.ca/"

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
        cell_value = sheet[f"{column_letter}{row_num}"].value

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

            driver.delete_all_cookies()
        else:
            print(f"Skipping row {row_num} because cell is empty")

except Exception as e:
    print("an unexpected error occured", str(e))
finally:
    driver.quit()
