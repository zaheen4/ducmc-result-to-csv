import os
import time
import gspread
import re
from google.oauth2.service_account import Credentials
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from selenium.common.exceptions import NoSuchElementException, TimeoutException
# from selenium.webdriver.chrome.options import Options

# --- CONFIGURATION ---
# All user-configurable settings are in this section.

# --- Google Sheets Configuration ---
GOOGLE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1aZWYV-V_YjdA_QRLNv5-vDx9YEWdaL2Is9RbK4ObLzw/edit?usp=drive_link'
CREDENTIALS_FILE = 'credentials.json'

# Defines the permissions the script requests from the Google API.
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
]

# Specifies the target worksheet within the Google Sheet.
WORKSHEET_NAME = 'PerCourse_L1T1'
# WORKSHEET_NAME = 'PerCourse_L1T2'
# WORKSHEET_NAME = 'PerCourse_L2T1'
# WORKSHEET_NAME = 'PerCourse_L2T2'
# WORKSHEET_NAME = 'PerCourse_L3T1'
# WORKSHEET_NAME = 'PerCourse_L3T2'


# --- Web Scraping Configuration ---
# Form data used to submit the search query on the results website.
FORM_DATA = {
    "program": "B.Sc. in Computer Science and Engineering",
    "session": "2021-2022",
    "exam": "B.Sc. in Computer Science and Engineering 1st year 1st Semester Examination of 2022"
    # "exam": "B.Sc. in Computer Science and Engineering 1st year 2nd Semester Examination of 2022"
    # "exam": "B.Sc. in Computer Science and Engineering 2nd year 1st Semester Examination of 2023"
    # "exam": "B.Sc. in Computer Science and Engineering 2nd year 2nd Semester Examination of 2023"
    # "exam": ""
    # "exam": ""
}
URL = 'http://cmc.du.ac.bd/result.php'

# Defines the range of registration numbers to scrape.
# Note: 710 - 813 is the full registration range for CSE-03.
START_REGI = 710
END_REGI = 813


def sanitize_text(text):
    """
    Cleans and standardizes text for reliable matching.
    Converts text to lowercase and removes all non-alphanumeric characters.
    Example: "CSE-1101" becomes "cse1101".
    """
    return re.sub(r'[^a-z0-9]', '', text.lower())


def parse_result_html(html_content):
    """
    Parses the HTML content of a student's result page to extract key information.

    Args:
        html_content (str): The raw HTML source of the result page.

    Returns:
        dict: A dictionary containing the student's name, registration, GPA, CGPA,
              failed subjects, and a list of courses with grades.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    student_data = {}

    # Extract basic student information (Name, Registration Number).
    info_rows = soup.select('div#exam_result > div.row > div.col-12 > table.table-bordered > tbody > tr')
    for row in info_rows:
        headers = row.find_all('th')
        if headers and len(headers) == 1:
            header_text = headers[0].get_text(strip=True)
            if "Student's Name" in header_text:
                student_data['Name'] = row.find('td').get_text(strip=True)
            elif "Registration" in header_text:
                student_data['Reg'] = row.find('td').get_text(strip=True)

    # Locate the result summary section to extract GPA, CGPA, and failed subjects.
    result_div = soup.find('div', style=lambda v: v and 'text-align: center' in v)

    # Initialize result fields to ensure they exist in the final dictionary.
    student_data['GPA'] = ''
    student_data['CGPA'] = ''
    student_data['Fail Subs'] = ''

    if result_div:
        result_text = result_div.get_text(separator=' ', strip=True)

        # Use regular expressions to reliably extract GPA and CGPA values.
        gpa_match = re.search(r"GPA:\s*([\d.]+)", result_text)
        if gpa_match:
            student_data['GPA'] = gpa_match.group(1)

        cgpa_match = re.search(r"CGPA:\s*([\d.]+)", result_text)
        if cgpa_match:
            student_data['CGPA'] = cgpa_match.group(1)

        # Find all course codes in the summary to identify failed or retake subjects.
        all_codes = re.findall(r'[A-Z]{2,}-\d{4}', result_div.get_text(separator=','))
        student_data['Fail Subs'] = ', '.join(all_codes) if all_codes else ''

    # Extract grades for each individual course from the grades table.
    student_data['courses'] = []
    course_table = soup.select_one('th table[width="100%"]')
    if course_table:
        # Skip the header row of the course table.
        for row in course_table.find_all('tr')[1:]:
            cols = row.find_all('td')
            if len(cols) == 5:
                course_name = cols[2].get_text(strip=True)
                grade_point = cols[4].get_text(strip=True)
                student_data['courses'].append({
                    "name": course_name,
                    "grade": grade_point if grade_point else '0.00' # Default to '0.00' if grade is missing.
                })
    return student_data


def main():
    """Main execution function to run the scraper and update the sheet."""
    print("Authenticating with Google Sheets...")
    try:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url(GOOGLE_SHEET_URL)
        worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
        print(f"✅ Successfully connected to '{spreadsheet.title}' and selected worksheet '{worksheet.title}'.")
    except Exception as e:
        print(f"FATAL ERROR: Could not connect to Google Sheets. {e}")
        return

    print("Fetching sheet data for comparison...")
    all_sheet_values = worksheet.get_all_values()
    sheet_headers = all_sheet_values[0]

    # Dynamically find the column indices for required fields.
    try:
        reg_col_index = sheet_headers.index('Reg. No.')
        gpa_col_index = sheet_headers.index('GPA')
        cgpa_col_index = sheet_headers.index('CGPA')
        retake_col_index = sheet_headers.index('Retake Courses')
    except ValueError as e:
        print(f"FATAL ERROR: A required column header was not found in Row 1 of your sheet: {e}")
        return

    # Create a list of all registration numbers currently in the sheet for quick lookups.
    all_reg_numbers_in_sheet = [row[reg_col_index] for row in all_sheet_values[1:]]

    # Build a map from sanitized course names to their column index for efficient updates.
    course_name_map = {}
    for i, header in enumerate(sheet_headers):
        if header.strip():
            name_part = header.split('\n')[0].strip()
            sanitized_name = sanitize_text(name_part)
            course_name_map[sanitized_name] = i

    # Initialize the Selenium WebDriver in headless mode (no visible browser window).
    options = Options()
    options.add_argument("--headless")
    driver = webdriver.Firefox(options=options)
    print("WebDriver initiated with Firefox in headless mode.")

    # --- Main processing loop for each registration number ---
    for regi_num in range(START_REGI, END_REGI + 1):
        current_reg = str(regi_num)
        print(f"\n--- Processing Registration No.: {current_reg} ---")
        try:
            driver.get(URL)
            wait = WebDriverWait(driver, 15)

            # Fill out the web form with the specified criteria.
            Select(wait.until(EC.presence_of_element_located((By.ID, 'pro_id')))).select_by_visible_text(FORM_DATA['program'])
            Select(wait.until(EC.presence_of_element_located((By.ID, 'sess_id')))).select_by_visible_text(FORM_DATA['session'])
            exam_xpath = f"//select[@id='exam_id']/option[text()='{FORM_DATA['exam']}']"
            wait.until(EC.element_to_be_clickable((By.XPATH, exam_xpath)))
            Select(driver.find_element(By.ID, 'exam_id')).select_by_visible_text(FORM_DATA['exam'])
            driver.find_element(By.ID, 'reg_no').send_keys(current_reg)
            driver.find_element(By.XPATH, "//button[text()='Submit']").click()

            # Wait for the result page to load before parsing.
            heading_xpath = "//h3[contains(text(), 'Result')]"
            wait.until(EC.presence_of_element_located((By.XPATH, heading_xpath)))
            time.sleep(0.5) # Brief pause for full page render.

            parsed_data = parse_result_html(driver.page_source)

            if not parsed_data or 'Reg' not in parsed_data:
                print(f"Result not found or page is invalid for Reg No: {current_reg}.")
                continue

            scraped_reg = parsed_data['Reg']
            print(f"Found result for Reg No: {scraped_reg} (Name: {parsed_data.get('Name', 'N/A')})")

            # Match the scraped registration number to its row in the Google Sheet.
            try:
                target_list_index = all_reg_numbers_in_sheet.index(scraped_reg)
                target_row_num = target_list_index + 2 # Adjust for 0-indexing and header row.
                existing_row_data = all_sheet_values[target_list_index + 1]
            except ValueError:
                print(f"Could not find registration '{scraped_reg}' in the sheet. Skipping.")
                continue

            print(f"Found registration on row {target_row_num}. Checking for empty cells before writing...")
            update_requests = []

            # --- Prepare updates only for empty cells to avoid overwriting existing data. ---
            if not existing_row_data[gpa_col_index] and parsed_data.get('GPA'):
                update_requests.append({'range': f'E{target_row_num}', 'values': [[parsed_data.get('GPA')]]})
            if not existing_row_data[cgpa_col_index] and parsed_data.get('CGPA'):
                update_requests.append({'range': f'F{target_row_num}', 'values': [[parsed_data.get('CGPA')]]})

            scraped_fail_subs = parsed_data.get('Fail Subs')
            if scraped_fail_subs and not existing_row_data[retake_col_index]:
                retake_col_letter = gspread.utils.rowcol_to_a1(1, retake_col_index + 1).rstrip('1')
                update_requests.append({'range': f'{retake_col_letter}{target_row_num}', 'values': [[scraped_fail_subs]]})

            for course in parsed_data.get('courses', []):
                scraped_name_sanitized = sanitize_text(course['name'])
                if scraped_name_sanitized in course_name_map:
                    col_index = course_name_map[scraped_name_sanitized]
                    if not existing_row_data[col_index]:
                        col_letter = gspread.utils.rowcol_to_a1(1, col_index + 1).rstrip('1')
                        update_requests.append({'range': f'{col_letter}{target_row_num}', 'values': [[course['grade']]]})

            # Perform a batch update to the sheet if there is new data to write.
            if update_requests:
                worksheet.batch_update(update_requests)
                print(f"✅ Successfully wrote {len(update_requests)} new value(s) to the Google Sheet.")
            else:
                print("No empty cells to update. All data for this student is already present.")

        except TimeoutException:
            print(f"No result found for {current_reg} (timeout).")
        except Exception as e:
            print(f"An unexpected error occurred for {current_reg}: {e}")

    driver.quit()
    print("\n--- All registration numbers processed. WebDriver closed. ---")

    # --- Post-processing: Convert and format GPA/CGPA columns ---
    print("\n--- Converting and formatting GPA/CGPA columns... ---")
    try:
        ranges_to_process = ["E3:E", "F3:F"]
        data_from_ranges = worksheet.batch_get(ranges_to_process, value_render_option='UNFORMATTED_VALUE')
        update_payload = []

        # Convert fetched string values to numbers where possible.
        for i, values_list in enumerate(data_from_ranges):
            range_name = ranges_to_process[i]
            converted_values = []
            for row in values_list:
                cell_value = row[0] if row else ''
                try:
                    numeric_value = float(cell_value)
                    converted_values.append([numeric_value])
                except (ValueError, TypeError):
                    converted_values.append([cell_value]) # Keep non-numeric values as is.

            # Prepare the data payload for the batch update.
            update_payload.append({
                'range': range_name,
                'values': converted_values
            })

        # Write the converted numbers back to the sheet.
        if update_payload:
            worksheet.batch_update(update_payload, value_input_option='USER_ENTERED')
            print("Step 1/2: Successfully converted any existing text values to numbers.")

        # Apply number formatting to the GPA/CGPA columns for consistency.
        worksheet.format(ranges_to_process, {
            "numberFormat": {
                "type": "NUMBER",
                "pattern": "0.00"
            }
        })
        print("Step 2/2: ✅ Successfully applied number formatting to GPA and CGPA columns.")

    except Exception as e:
        print(f"⚠️ Could not apply formatting. Error: {e}")


if __name__ == '__main__':
    main()
