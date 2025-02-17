import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from webdriver_manager.chrome import ChromeDriverManager

# Load Excel file
excel_file = "PY-6(Complete).xlsx"
df = pd.read_excel(excel_file)


# Ensure the email column exists
if "Email" not in df.columns:
    raise ValueError("Email column not found in the Excel file!")

# Set up Chrome options
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.prompt_for_download": False,
    "plugins.always_open_pdf_externally": True,
})

# Initialize the WebDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

def login():
    driver.get("https://training.edge.gov.bd/login")
    time.sleep(3)
    driver.find_element(By.ID, "email").send_keys("kamrulhasan78666@gmail.com")
    driver.find_element(By.ID, "password").send_keys("Barishel@2024")
    driver.find_element(By.ID, "password").send_keys(Keys.RETURN)
    time.sleep(3)

def navigate_to_batches():
    driver.get("https://training.edge.gov.bd/student-training/batches/5488")
    time.sleep(3)

    try:
        # Select "100" from the dropdown menu
        dropdown = driver.find_element(By.ID, "dt-length-0")
        dropdown.click()
        time.sleep(1)
        dropdown.find_element(By.XPATH, "//option[@value='100']").click()
        print("✅ Selected '100' trainees per page.")
        time.sleep(3)  # Wait for table to update
    except NoSuchElementException:
        print("⚠️ Could not find the trainee count dropdown!")

def get_all_trainee_ids():
    trainee_ids = []
    while True:
        rows = driver.find_elements(By.XPATH, "//table[@id='traineeTable']//tbody//tr")
        for row in rows:
            try:
                trainee_link = row.find_element(By.XPATH, ".//a[contains(@href, '/trainees/')]")
                trainee_id = trainee_link.get_attribute("href").split("/")[-1]
                trainee_ids.append(trainee_id)
            except Exception as e:
                print(f"Error extracting trainee ID: {e}")
        
        try:
            next_button = driver.find_element(By.ID, "traineeTable_next")
            if "disabled" in next_button.get_attribute("class"):
                break
            next_button.click()
            time.sleep(3)
        except Exception:
            break
    return trainee_ids

def navigate_to_trainee_page(trainee_id):
    driver.get(f"https://training.edge.gov.bd/student-training/batches/5672/trainees/{trainee_id}")
    time.sleep(3)

def get_trainee_email():
    try:
        email_element = driver.find_element(By.XPATH, "//label[text()='Email Address']/following-sibling::div")
        return email_element.text.strip()
    except NoSuchElementException:
        return None

def click_assessment_button():
    try:
        # Try to click "Add Assessment and Digital Profile Information" button
        button = driver.find_element(By.XPATH, "//button[contains(text(), 'Add Assessment and Digital Profile Information')]")
        button.click()
    except NoSuchElementException:
        try:
            # If not found, click "Edit Assessment and Digital Profile Information" button
            button = driver.find_element(By.XPATH, "//button[contains(text(), 'Edit Assessment and Digital Profile Information')]")
            button.click()
        except NoSuchElementException:
            print("Assessment button not found!")

    time.sleep(3)

def fill_input_field(field_id, value):
    try:
        # Wait until the input field is clickable before interacting
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, field_id)))
        input_box = driver.find_element(By.ID, field_id)
        input_box.clear()
        input_box.send_keys(str(value))  # Ensure the value is converted to string
        print(f"✅ Filled {field_id} with value: {value}")  # Log the value being entered
    except NoSuchElementException:
        print(f"⚠️ Input field {field_id} not found!")
    except Exception as e:
        print(f"Error filling field {field_id}: {e}")

def click_save_button():
    try:
        save_button = driver.find_element(By.XPATH, "//button[@id='submitModal' and contains(text(), 'Save')]")
        save_button.click()
    except NoSuchElementException:
        print("Save button not found!")

def click_update_button():
    try:
        update_button = driver.find_element(By.XPATH, "//button[@id='submitModal' and contains(text(), 'Update')]")
        update_button.click()
    except NoSuchElementException:
        print("Update button not found!")

def handle_alert():
    try:
        alert = driver.switch_to.alert
        alert.accept()
        print("✅ Alert accepted!")
    except:
        print("No alert found.")

def display_trainee_information(trainee_data):
    print(f"📌 Roll: {trainee_data['Roll']}")
    print(f"📌 Name: {trainee_data['Trainee Name']}")
    print(f"📌 Phone: {trainee_data['Phone']}")
    print(f"📌 Email: {trainee_data['Email']}")
    print(f"📌 LinkedIn: {trainee_data['LinkedIn Account']}")
    print(f"📌 GitHub: {trainee_data['Github Link(Project)']}")
    print(f"📌 Fiverr/Upwork: {trainee_data.get('Fiverr  or Upwork  Account Link', 'N/A')}")
    
def process_all_trainees():
    trainee_ids = get_all_trainee_ids()
    for trainee_id in trainee_ids:
        navigate_to_trainee_page(trainee_id)
        email = get_trainee_email()
        
        if email:
            matching_row = df[df["Email"].str.strip().str.lower() == email.lower()]
            if not matching_row.empty:
                trainee_data = matching_row.iloc[0]

                # Ensure values are strings and handle NaN cases
                linkedin = str(trainee_data.get("LinkedIn Account", "")).strip()
                github = str(trainee_data.get("Github Link(Project)", "")).strip()
                fiverr = str(trainee_data.get("Fiverr  or Upwork  Account Link", "")).strip()

                # Skip trainee if any required value is missing
                if not linkedin or not github or not fiverr:
                    print(f"⏭️ Skipping Trainee ID {trainee_id} (Missing LinkedIn/GitHub/Fiverr)")
                    continue

                print(f"✅ Processing Trainee ID: {trainee_id}, Email: {email}")

                click_assessment_button()
                
                # Fetch the marks values from the Excel sheet
                attendance = trainee_data.get("Attendance (10%)", "")
                quiz = trainee_data.get("Quiz (10%)", "")
                midterm = trainee_data.get("Mid Term (20%)", "")
                project = trainee_data.get("Project (25%)", "")
                final_evaluation = trainee_data.get("Final Evaluation (25%)", "")

                # Fill the input fields with the corresponding values
                fill_input_field("attendance_percentage", attendance)
                fill_input_field("quiz_assessment_marks", quiz)
                fill_input_field("midterm_assessment_marks", midterm)
                fill_input_field("project_assessment_marks", project)
                fill_input_field("final_assessment_marks", final_evaluation)

                # Now fill the LinkedIn, GitHub, Fiverr fields
                fill_input_field("linkedin_profile", linkedin)
                fill_input_field("link_of_projects_repository", github)
                fill_input_field("link_of_freelancing_profile", fiverr)

                try:
                    if driver.find_elements(By.XPATH, "//button[contains(text(), 'Edit Assessment and Digital Profile Information')]"):
                        click_update_button()
                    else:
                        click_save_button()
                    time.sleep(2)
                    handle_alert()
                except Exception as e:
                    print(f"Error determining button action: {e}")
            else:
                print(f"Trainee ID: {trainee_id}, Email: {email} - No match found in Excel.")

try:
    login()
    navigate_to_batches()
    process_all_trainees()
finally:
    driver.quit()
