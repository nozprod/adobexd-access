from selenium import webdriver
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, ElementClickInterceptedException
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

# Main function
def automate(email, username, password, csv_path, group):
    # Load csv
    df = pd.read_csv(csv_path, header=None)
    links = df[0].tolist()
    total_links = len(links)

    # Initialize driver using webdriver_manager
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    # List to store status
    report = []

    # Connect to Adobe.com
    max_attempts = 3
    attempt = 0
    logged = False

    while attempt < max_attempts and not logged:
        try:
            driver.get(links[0])
            email_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'EmailPage-EmailField')))
            email_field.send_keys(email)
            email_field.send_keys(Keys.RETURN)

            user_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'username')))
            user_field.send_keys(username)

            password_field = driver.find_element(By.ID, 'password')
            password_field.send_keys(password)
            password_field.send_keys(Keys.RETURN)

            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.inviteButton-I4RMS')))

            logged = True

        except TimeoutException:
            print(f"\033[91mLogin failed, trying again\033[0m")
            attempt += 1

    if logged:
        print(f"\033[92mLogin successful\033[0m")
        for index, link in enumerate(links[0:]):
            try:
                # Open link
                driver.get(link)

                status = {"Link": link, "Status": "NOK"}  # Default status is NOK
                performed = False

                try:
                    # Click on invite button
                    invite_button = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, '.inviteButton-I4RMS'))
                    )
                    invite_button.click()

                    # Fill in a group
                    input_field = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, 'ccx-ss-flex-input-textarea'))
                    )
                    input_field.send_keys(group)
                    wait = WebDriverWait(driver, 10)
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
                    print(f"\033[91m{index + 1}/{total_links}.\033[0m No invite button found")

                # Ask access if needed
                if not performed:
                    try:
                        # Try to find the request access button
                        access_button = WebDriverWait(driver, 10).until(
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
        print(f"\033[91mLogin failed after many attempts\033[0m")

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
email_entry.insert(0, "")

tk.Label(root, text="Username:").grid(row=1)
username_entry = tk.Entry(root)
username_entry.grid(row=1, column=1)
username_entry.insert(0,"")

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
group_entry.insert(0, "")

# Execution button
tk.Button(root, text="Execute", command=submit).grid(row=5, columnspan=2)

# Launch GUI
root.mainloop()
