from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import time
import csv
from selenium.webdriver.firefox.options import Options



######## edit form data

start_regi = 710     # CSE(21-22) for now
end_regi = 813


regi = start_regi
program = "B.Sc. in Computer Science and Engineering"
session = "2021-2022"
exam = "B.Sc. in Computer Science and Engineering 1st year 2nd Semester Examination of 2022"

######

file_path = "Output/ResultFromWebsite.txt"
csv_file_path = "Output/Results.csv"
delim = ";"  # defines the seperator used for CSV file


# options = Options()
# options.add_argument("--width=576")  # Set window width
# options.add_argument("--height=324")  # Set window height options=options

driver = webdriver.Firefox()
driver.get("https://ducmc.com/result.php")


# Prevents the loop from running infinitely
iteration_count = 0
max_iteration = 240


# find CG of every student in dept
while (regi <= end_regi) and (iteration_count < max_iteration):
   try:
      # verifies if the site loaded 
      assert "DUCMC" in driver.title 

      #time.sleep(1)

      # inputs Registration No
      registration_no = driver.find_element(By.ID, "reg_no")
      registration_no.clear()
      registration_no.send_keys(regi)
      

      # filling Program/Course
      program_name = driver.find_element(By.ID, "pro_id")
      Select(program_name).select_by_visible_text(program)

      #time.sleep(1) 

      # filling Session
      session_year = driver.find_element(By.ID, "sess_id")
      Select(session_year).select_by_visible_text(session)

      # filling Exam name
      exam_name = driver.find_element(By.ID, "exam_id")
      Select(exam_name).select_by_visible_text(exam)

      #time.sleep(1)

      # clicking submit
      submit_button = driver.find_element(By.CLASS_NAME, "btn.btn-primary.btn-block")
      submit_button.click()

      #time.sleep(2)

      # finding GPA data
      # result_data = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div/div/div[2]/div/div[2]/table/tbody/tr[11]/td/div/small[2]')
      try:
         result_data = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div/div/div[2]/div/div[2]/table/tbody/tr[11]/td/div')
         result_text = result_data.text
         print(f"\n{result_text}")

      except NoSuchElementException:
         # indicates absent [or the Result page didn't load]
         result_text = "Absent"



   ####### TXT file part (Optional)
      # with open(file_path, "w") as file:
      #    file.write(result_text)

      # with open(file_path, "r") as file:
      #    content = file.read()

      content = result_text

      ## Formatting
      # Replace all newline characters with custom delimeter for CSV
      content = content.replace("\n", delim)
      #content = content.replace(":", delim)


      ## Formatting
      # Add an extra cell for failed subjects if it doesn't exist already
      semicolon_count = content.count(";")

      if semicolon_count == 2:
         parts = content.split(";")
         
         # Insert an extra semicolon after the 2nd semicolon
         modified_content = ";".join(parts[:1] + [""] + parts[1:])
      else:
         # Keep the content unchanged
         modified_content = content
      


      # kinda acts like a log file for real time viewing (current row)
      with open(file_path, "w") as file:
         file.write(f"{regi}{delim}{modified_content}")

      
      
   ##### CSV Part
      # Open the file in append mode then write the content
      with open(csv_file_path, mode="a", newline="", encoding="utf-8") as csvfile:
         # writer = csv.writer(csvfile, delimiter=";")
         # writer.writerow(data)
         csvfile.write(f"{regi}{delim}{modified_content}\n")   # Appending directly as the text is already formatted 

   ##### CSV end

   except Exception as e:
      print(f"Error for regi number {regi}: {e}") 


   regi += 1 
   iteration_count += 1

   # refresh the current webpage
   driver.refresh()

driver.close()