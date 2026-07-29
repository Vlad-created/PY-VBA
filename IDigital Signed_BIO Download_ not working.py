from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from datetime import date, timedelta
import time

# -----------------------------
# Calculate last month dates
# -----------------------------
today = date.today()
first_day_this_month = today.replace(day=1)
last_day_last_month = first_day_this_month - timedelta(days=1)
first_day_last_month = last_day_last_month.replace(day=1)

start_date_str = first_day_last_month.strftime("%m/%d/%Y")
end_date_str = last_day_last_month.strftime("%m/%d/%Y")

# -----------------------------
# Chrome with your real profile (bypasses CAPTCHA)
# -----------------------------
options = Options()
options.add_argument("user-data-dir=C:/Users/YOURUSER/AppData/Local/Google/Chrome/User Data")
options.add_argument("profile-directory=Default")

driver = webdriver.Chrome(options=options)

# -----------------------------
# Open Power BI RDL report
# -----------------------------
driver.get("https://app.powerbi.com/groups/me/rdlreports/9a948558-c4b5-41e7-a329-72458ace83aa")

wait = WebDriverWait(driver, 60)

# -----------------------------
# Switch into Power BI iframe
# -----------------------------
wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
driver.switch_to.frame(0)

# -----------------------------
# Fill Start Date
# -----------------------------
start_input = wait.until(
    EC.presence_of_element_located((By.XPATH, "//input[@aria-label='Start date']"))
)
start_input.clear()
start_input.send_keys(start_date_str)

# Fill End Date
end_input = wait.until(
    EC.presence_of_element_located((By.XPATH, "//input[@aria-label='End date']"))
)
end_input.clear()
end_input.send_keys(end_date_str)

# -----------------------------
# Click View report
# -----------------------------
view_button = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//button//span[text()='View report']"))
)
view_button.click()

# -----------------------------
# Wait for report to reload
# -----------------------------
driver.switch_to.default_content()
time.sleep(5)

wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
driver.switch_to.frame(0)

# -----------------------------
# Open Export Menu
# -----------------------------
export_button = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//button[contains(@title,'Export')]"))
)
export_button.click()

# -----------------------------
# Click Excel option
# -----------------------------
excel_option = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'Microsoft Excel (.xlsx)')]"))
)
excel_option.click()

# -----------------------------
# Wait for download
# -----------------------------
time.sleep(10)
driver.quit()
