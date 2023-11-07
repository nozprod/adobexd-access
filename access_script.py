from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import tkinter as tk
from tkinter import filedialog

import time
import pandas as pd

#Main function
def automate(email, username, password, csv_path, group):

    # Load csv
    df = pd.read_csv(csv_path, header=None)
    links = df[0].tolist()

    # Initialize driver
    driver = webdriver.Chrome()

    # Connect to Adobe.com
    driver.get(links[0])
    email_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'EmailPage-EmailField')))
    email_field.send_keys(email)
    email_field.send_keys(Keys.RETURN)

    user_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'username')))
    user_field.send_keys(username)

    password_field = driver.find_element(By.ID, 'password')
    password_field.send_keys(password)
    password_field.send_keys(Keys.RETURN)

    time.sleep(1)  # Stop for 1 second (needed for login acceptance)

    for lien in links[0:]:
        # Open link
        driver.get(lien)

        performed = False

        try:
            # Clic on invite button
            invite_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.inviteButton-I4RMS')))
            invite_button = driver.find_element(By.CSS_SELECTOR, '.inviteButton-I4RMS')
            invite_button.click()

            # # Fill in a mail
            # input_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'ccx-ss-flex-input-textarea')))
            # input_field = driver.find_element(By.ID, 'ccx-ss-flex-input-textarea')
            # input_field.send_keys('valentin.cazin@stellantis.com')
            # input_field.send_keys(Keys.RETURN)

            # Fill in a group 
            input_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'ccx-ss-flex-input-textarea')))
            input_field = driver.find_element(By.ID, 'ccx-ss-flex-input-textarea')
            input_field.send_keys(group)
            wait = WebDriverWait(driver, 10)
            suggestions = wait.until(EC.visibility_of_element_located((By.ID, 'ccx-ss-suggestion-0')))
            suggestions.click()

            time.sleep(1)  # Stop for 1 second (needed for invite acceptance, generate error otherwise)

            # # Select invite mode
            # mode_dropdown = driver.find_element(By.ID, 'react-spectrum-2203')
            # mode_dropdown.click()
            # desired_mode = driver.find_element_by_xpath("//option[text()='Mode_d_invitation']")
            # desired_mode.click()

            # Execute invitation
            final_invite_button = driver.find_element(By.ID, 'ccx-ss-share-invite-send-btn')
            final_invite_button.click()

            performed = True

        except TimeoutException:
            print(f"No invite button found{lien}")

        #Ask access if needed
        if not performed:
            try:
                # Try to find the request access button
                access_button = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "[data-auto='requestAccessButton']"))
                )
                
                access_button.click()
                print(f"Access required for {lien}")

            except TimeoutException:
                print(f"Access already asked for {lien}")
                continue   

    # Close browser when done
    driver.quit()

#GUI
def submit():
    email = email_entry.get()
    username = username_entry.get()
    password = password_entry.get()
    csv_path = csv_entry.get()
    group = group_entry.get()
    
    automate(email, username, password, csv_path, group)

# Main frame
root = tk.Tk()
root.title("Adobe XD - Access script")

# Fields and labels
tk.Label(root, text="Email:").grid(row=0)
email_entry = tk.Entry(root)
email_entry.grid(row=0, column=1)
email_entry.insert(0, "@stellantis.com")  # Pre-filled

tk.Label(root, text="Username:").grid(row=1)
username_entry = tk.Entry(root)
username_entry.grid(row=1, column=1)
username_entry.insert(0, "U533621")  # Pre-filled /!\ TO BE REMOVED WHEN DISTRIBUTED

tk.Label(root, text="Password:").grid(row=2)
password_entry = tk.Entry(root, show='*')
password_entry.grid(row=2, column=1)
password_entry.insert(0, "vC180890")  # Pre-filled /!\ TO BE REMOVED WHEN DISTRIBUTED

tk.Label(root, text="CSV File:").grid(row=3)
csv_entry = tk.Entry(root)
csv_entry.grid(row=3, column=1)
tk.Button(root, text="Browse", command=lambda: csv_entry.insert(0, filedialog.askopenfilename())).grid(row=3, column=2)

tk.Label(root, text="Group:").grid(row=4)
group_entry = tk.Entry(root)
group_entry.grid(row=4, column=1)
group_entry.insert(0, "UX_Plateau_")  # Pre-filled /!\ NEXT STEP --> LOAD FROM A CSV

# Execution button
tk.Button(root, text="Execute", command=submit).grid(row=5, columnspan=2)

# Launch GUI
root.mainloop()