import sys
import os
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, ElementClickInterceptedException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

import tkinter as tk
from tkinter import filedialog

from datetime import datetime
import time
import pandas as pd


#Login check
def try_login(driver, email, username, password, links):

    max_attempts = 3
    attempt = 0

    while attempt < max_attempts:
        try:
            driver.get(links[0])
            email_field = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'EmailPage-EmailField')))
        except TimeoutException:
            print(f"\033[92mLogin successful\033[0m")
            return True

        try:
            email_field.send_keys(email)
            email_field.send_keys(Keys.RETURN)

            user_field = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'username')))
            user_field.send_keys(username)

            password_field = driver.find_element(By.ID, 'password')
            password_field.send_keys(password)
            password_field.send_keys(Keys.RETURN)

            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.inviteButton-I4RMS')))
            print(f"\033[92mLogin successful\033[0m")
            return True
        except TimeoutException:
            print(f"\033[91mLogin failed, trying again\033[0m")
            attempt += 1

    print(f"\033[91mLogin failed after many attempts\033[0m")
    return False


# Check requested access
def check_and_process_access_requests(driver, domains_allowed):
    try:
        access_requests = WebDriverWait(driver, 0.1).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.ccx-ss-request-card'))
        )

        for request in access_requests:
            email_element = request.find_element(By.CSS_SELECTOR, ".ccx-ss-request-card-second-line")
            email_text = email_element.text
            domain_of_email = email_text.split('@')[-1]

            if any(domain_of_email.endswith(domain_allowed) for domain_allowed in domains_allowed):
                accept_button = request.find_element(By.CSS_SELECTOR, ".ccx-ss-collaborators-list-request-accept-btn")
                accept_button.click()
                print(f"Access request from {email_text} accepted.")
            else:
                decline_button = request.find_element(By.CSS_SELECTOR, ".ccx-ss-collaborators-list-request-decline-btn")
                decline_button.click()
                print(f"Access request from {email_text} rejected.")
    except TimeoutException:
        pass

# Allowed domains
domains_allowed = ["@stellantis.com", "@external.stellantis.com"]


# Main function
def automate(email, username, password, csv_path, group):
    # Load csv
    df = pd.read_csv(csv_path, header=None)
    links = df[0].tolist()
    total_links = len(links)

    # Initialize driver using webdriver_manager
    # driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    # Initialize driver using local path
    if getattr(sys, 'frozen', False):
        chromedriver_path = os.path.join(sys._MEIPASS, 'chromedriver.exe')
    else:
        chromedriver_path = './chromedriver'

    s = Service(executable_path=chromedriver_path)
    driver = webdriver.Chrome(service=s)

    # List to store status
    report = []

    logged = try_login(driver, email, username, password, links)

    if logged:
        for index, link in enumerate(links[0:]):
            try:
                # Open link
                driver.get(link)

                status = {"Link": link, "Status": "NOK"}  # Default status is NOK
                performed = False

                max_attempts = 2
                attempt = 0

                while attempt < max_attempts and performed != True:
                    try:
                        # Click on invite button
                        invite_button = WebDriverWait(driver, 20).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, '.inviteButton-I4RMS'))
                        )
                        invite_button.click()

                        # Fill in a group
                        input_field = WebDriverWait(driver, 20).until(
                            EC.presence_of_element_located((By.ID, 'ccx-ss-flex-input-textarea'))
                        )

                        # Check access requests
                        check_and_process_access_requests(driver, domains_allowed)

                        input_field.send_keys(group)
                        wait = WebDriverWait(driver, 20)
                        suggestions = wait.until(EC.visibility_of_element_located((By.ID, 'ccx-ss-suggestion-0')))
                        suggestions.click()

                        time.sleep(1)  # Stop for 1 second (needed for invite acceptance)

                        # Execute invitation
                        final_invite_button = driver.find_element(By.ID, 'ccx-ss-share-invite-send-btn')
                        final_invite_button.click()

                        print(f"\033[92m{index + 1}/{total_links}.\033[0m Invite sent")
                        status["Status"] = "OK"
                        performed = True

                        # Add status to report
                        report.append(status)

                    except TimeoutException:
                        print(f"\033[91m{index + 1}/{total_links}.\033[0m No invite button found, trying again")
                        attempt += 1

                # Ask access if needed
                if not performed:
                    try:
                        # Try to find the request access button
                        access_button = WebDriverWait(driver, 20).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "[data-auto='requestAccessButton']"))
                        )

                        access_button.click()
                        print(f"\033[91m{index + 1}/{total_links}.\033[0m Access required")
                        status["Status"] = "NOK - Access required"

                        # Add status to report
                        report.append(status)

                    except TimeoutException:
                        print(f"\033[91m{index + 1}/{total_links}.\033[0m Access already asked")
                        status["Status"] = "NOK - Access required"

                        # Add status to report
                        report.append(status)
                        continue

            except StaleElementReferenceException:
                print(f"\033[91m{index + 1}/{total_links}.\033[0m StaleElementReference error, opening next link")
                status["Status"] = "NOK - Error"
                report.append(status)
                continue

            except ElementClickInterceptedException:
                print(f"\033[91m{index + 1}/{total_links}.\033[0m ElementClickIntercepted error, opening next link")
                status["Status"] = "NOK - Error"
                report.append(status)
                continue

            except Exception as e:
                print(f"\033[91m{index + 1}/{total_links}.\033[0m Unexpected error: {e}")
                status["Status"] = "NOK - Error"
                report.append(status)
                continue

        # Close browser when done
        driver.quit()

        # Generate report
        date_time_now = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        report_name = f"{date_time_now}_report.csv"

        df_report = pd.DataFrame(report)
        df_report.to_csv(report_name, index=False, encoding='utf-8-sig')
        print("\033[92mReport generated successfully\033[0m")

    else:
        pass

# GUI
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
