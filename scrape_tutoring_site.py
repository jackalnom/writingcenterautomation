from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import pandas as pd
import os
from dotenv import load_dotenv

load_dotenv()

tutoring_site_url = os.getenv('tutoring_site_url')
tutoring_site_username = os.getenv('tutoring_site_username')
tutoring_site_password = os.getenv('tutoring_site_password')

driver = webdriver.Chrome("chromedriver")
driver.maximize_window()

driver.get(f"{tutoring_site_url}/index.php?logout=YES")
assert 'WCONLINE' in driver.title

# find username/email field and send the username itself to the input field
driver.find_element("id", "username").send_keys(tutoring_site_username)
# find password input field and insert password as well
driver.find_element("id", "password").send_keys(tutoring_site_password)
# click login button
driver.find_element("name", "submit").click()

driver.get(f"{tutoring_site_url}/admin_reportmlr.php?reset=1")

select = Select(driver.find_element("id", "climit"))

select.select_by_value("orphan")
# sdate
# edate
submit_button = driver.find_element("name", "submit")
submit_button.click()

rows = driver.find_elements(By.CLASS_NAME, "mb-3")

cleaned = []

# I start with the third element here because that is when the student data starts
for i in range(3, len(rows)):
    # Data cleanup
    date, time, type, student, tutor, schedule, options = rows[i].text.split("\n")
    tutor = tutor.split(":")[1].strip()
    schedule = schedule.split(":")[1].strip()

    print(f"Date: {date} Time: {time} Tutor: {tutor} Student: {student} Schedule: {schedule}")
    cleaned.append([date, time, tutor, student, schedule])

pd.DataFrame(cleaned, columns=['date', 'time', 'tutor', 'student', 'room']).to_csv("orphans.csv")

driver.quit()