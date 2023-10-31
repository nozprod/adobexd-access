from selenium import webdriver
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

    # Charger le fichier Excel
    df = pd.read_csv(csv_path, header=None)
    links = df[0].tolist()

    # Initialiser le driver
    driver = webdriver.Chrome()

    # Se connecter
    driver.get(links[0])
    email_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'EmailPage-EmailField')))
    #email_field = driver.find_element(By.ID, 'EmailPage-EmailField')
    email_field.send_keys(email)
    email_field.send_keys(Keys.RETURN)

    user_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'username')))
    #user_field = driver.find_element(By.ID, 'username')
    user_field.send_keys(username)

    password_field = driver.find_element(By.ID, 'password')
    password_field.send_keys(password)
    password_field.send_keys(Keys.RETURN)

    time.sleep(1)  # pause le script pendant 1 seconde (nécessaire pour prise en compte du login)

    for lien in links[0:]:
        # Ouvrir le lien
        driver.get(lien)

        #Ask access if needed
        try:
            # Tentative de localisation du bouton "Demander l'accès"
            access_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "[data-auto='requestAccessButton']"))  # Remplacez 'access_button_id' par le bon identifiant ou sélecteur
            )
            
            # Si le bouton est trouvé, cliquez dessus et continuez avec la prochaine itération de la boucle
            access_button.click()
            print(f"Access required for {lien}")
            continue  # Passe au prochain lien

        except:
            # Si le bouton n'est pas trouvé, le script continue normalement
            pass
        
        # Cliquez sur le bouton pour inviter
        invite_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.inviteButton-I4RMS')))
        invite_button = driver.find_element(By.CSS_SELECTOR, '.inviteButton-I4RMS')
        invite_button.click()

        # # Renseigner un email
        # input_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'ccx-ss-flex-input-textarea')))
        # input_field = driver.find_element(By.ID, 'ccx-ss-flex-input-textarea')
        # input_field.send_keys('valentin.cazin@stellantis.com')
        # input_field.send_keys(Keys.RETURN)

        # Renseigner un groupe
        input_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'ccx-ss-flex-input-textarea')))
        input_field = driver.find_element(By.ID, 'ccx-ss-flex-input-textarea')
        input_field.send_keys(group)
        wait = WebDriverWait(driver, 10)
        suggestions = wait.until(EC.visibility_of_element_located((By.ID, 'ccx-ss-suggestion-0')))
        suggestions.click()

        time.sleep(1)  # pause le script pendant 1 seconde (nécessaire pour prise en compte de la saisie, génère une erreur sinon)

        # # Sélectionner le mode d'invitation
        # mode_dropdown = driver.find_element(By.ID, 'react-spectrum-2203')
        # mode_dropdown.click()
        # desired_mode = driver.find_element_by_xpath("//option[text()='Mode_d_invitation']") # Remplacez 'Mode_d_invitation'
        # desired_mode.click()

        # Cliquez sur "inviter"
        final_invite_button = driver.find_element(By.ID, 'ccx-ss-share-invite-send-btn')
        final_invite_button.click()

    # Fermer le navigateur à la fin
    driver.quit()

#GUI
def submit():
    email = email_entry.get()
    username = username_entry.get()
    password = password_entry.get()
    csv_path = csv_entry.get()
    group = group_entry.get()
    
    automate(email, username, password, csv_path, group)

# Créer la fenêtre principale
root = tk.Tk()
root.title("Adobe XD - Access script")

# Ajouter des champs de texte et des étiquettes
tk.Label(root, text="Email:").grid(row=0)
email_entry = tk.Entry(root)
email_entry.grid(row=0, column=1)
email_entry.insert(0, "@stellantis.com")  # Pré-remplir le champ

tk.Label(root, text="Username:").grid(row=1)
username_entry = tk.Entry(root)
username_entry.grid(row=1, column=1)
username_entry.insert(0, "U533621")  # Pré-remplir le champ /!\ À supprimer lors de la distribution

tk.Label(root, text="Password:").grid(row=2)
password_entry = tk.Entry(root, show='*')
password_entry.grid(row=2, column=1)
password_entry.insert(0, "vC180890")  # Pré-remplir le champ /!\ À supprimer lors de la distribution

tk.Label(root, text="CSV File:").grid(row=3)
csv_entry = tk.Entry(root)
csv_entry.grid(row=3, column=1)
tk.Button(root, text="Browse", command=lambda: csv_entry.insert(0, filedialog.askopenfilename())).grid(row=3, column=2)

tk.Label(root, text="Group:").grid(row=4)
group_entry = tk.Entry(root)
group_entry.grid(row=4, column=1)
group_entry.insert(0, "UX_Plateau_")  # Pré-remplir le champ

# Ajouter un bouton de soumission
tk.Button(root, text="Execute", command=submit).grid(row=5, columnspan=2)

# Exécuter l'interface utilisateur
root.mainloop()