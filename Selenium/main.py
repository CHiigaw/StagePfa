import time
import os
import zipfile
import re
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Chemin vers votre chromedriver
chrome_driver_path = 'C:/Users/DELL/Documents/Selenium/chromedriver.exe'

# URL du site web
url = 'https://www.marchespublics.gov.ma/index.php?page=entreprise.EntrepriseAdvancedSearch&searchAnnCons'

# Répertoire de téléchargement
download_dir = 'C:\\Users\\DELL\\YUKI'
data_dir = 'C:\\Users\\DELL\\DATAEX'

# Créer le répertoire DATA s'il n'existe pas
if not os.path.exists(data_dir):
    os.makedirs(data_dir)

# Options Chrome
chrome_options = Options()
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_experimental_option('prefs', {
    'download.default_directory': download_dir,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'safebrowsing.enabled': True
})

# Créer une nouvelle instance du driver Chrome
driver = webdriver.Chrome(service=Service(executable_path=chrome_driver_path), options=chrome_options)

try:
    # Ouvrir le site web
    driver.get(url)

    # Sélectionner "Mode de passation"
    mode_de_passation_select = Select(driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_AdvancedSearch_procedureType'))
    mode_de_passation_select.select_by_value("1")

    # Sélectionner "Catégorie principale"
    categorie_principal_select = Select(driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_AdvancedSearch_categorie'))
    categorie_principal_select.select_by_value("2")

    # Cliquer sur 'Détails' pour 'Lieu d'exécution'
    details_link = driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_AdvancedSearch_linkLieuExe1')
    driver.execute_script("arguments[0].click();", details_link)

    # Basculer vers la nouvelle fenêtre qui s'ouvre
    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)

    # Attendre que le bouton radio "Sélectionner par province(s)" soit disponible et cliquer dessus
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'ctl0_CONTENU_PAGE_repeaterGeoN0_ctl0_selectiongeoN0Select'))
    )
    driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_repeaterGeoN0_ctl0_selectiongeoN0Select').click()

    # Attendre que le menu déroulant Rabat-Salé-Kénitra soit disponible et cliquer pour l'ouvrir
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[contains(normalize-space(text()), 'Rabat-Salé-Kénitra')]"))
    )
    driver.find_element(By.XPATH, "//div[contains(normalize-space(text()), 'Rabat-Salé-Kénitra')]").click()

    # Attendre que le menu déroulant Casablanca-Settat soit disponible et cliquer pour l'ouvrir
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[contains(normalize-space(text()), 'Casablanca-Settat')]"))
    )
    driver.find_element(By.XPATH, "//div[contains(normalize-space(text()), 'Casablanca-Settat')]").click()

    # Attendre que la case "Tous" de Rabat-Salé-Kénitra soit disponible et la cocher
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'ctl0_CONTENU_PAGE_repeaterGeoN0_ctl0_repeaterGeoN1_ctl4_region1'))
    )
    tous_rabat_checkbox = driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_repeaterGeoN0_ctl0_repeaterGeoN1_ctl4_region1')
    driver.execute_script("arguments[0].scrollIntoView();", tous_rabat_checkbox)
    tous_rabat_checkbox.click()

    # Attendre que la case "Tous" de Casablanca-Settat soit disponible et la cocher
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'ctl0_CONTENU_PAGE_repeaterGeoN0_ctl0_repeaterGeoN1_ctl1_region2'))
    )
    tous_casablanca_checkbox = driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_repeaterGeoN0_ctl0_repeaterGeoN1_ctl1_region2')
    driver.execute_script("arguments[0].scrollIntoView();", tous_casablanca_checkbox)
    tous_casablanca_checkbox.click()

    # Cliquer sur le bouton "Valider"
    valider_button = driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_validateButton')
    valider_button.click()

    # Attendre d'être redirigé vers la première fenêtre
    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(1))
    driver.switch_to.window(driver.window_handles[0])

    # Cliquer sur 'Définir' pour les domaines d'activité
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'ctl0_CONTENU_PAGE_AdvancedSearch_domaineActivite_linkDisplay'))
    )
    definir_button = driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_AdvancedSearch_domaineActivite_linkDisplay')
    driver.execute_script("arguments[0].click();", definir_button)

    # Basculer vers la nouvelle fenêtre qui s'ouvre
    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)

    # Attendre que les cases à cocher soient disponibles et cliquer dessus

    # Cocher "Matériel, mobilier et fournitures de bureau"
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'ctl0_CONTENU_PAGE_repeaterCategorie_ctl1_repeaterSousCategorie_ctl12_idSection'))
    )
    bureau_checkbox = driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_repeaterCategorie_ctl1_repeaterSousCategorie_ctl12_idSection')
    driver.execute_script("arguments[0].scrollIntoView();", bureau_checkbox)
    bureau_checkbox.click()

    # Cocher "Documentation, manuels, fournitures scolaires et d’enseignement"
    documentation_checkbox = driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_repeaterCategorie_ctl1_repeaterSousCategorie_ctl10_idSection')
    driver.execute_script("arguments[0].scrollIntoView();", documentation_checkbox)
    documentation_checkbox.click()

    # Cliquer sur le bouton "Valider"
    valider_button_domaine = driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_validateButton')
    valider_button_domaine.click()

    # Revenir à la fenêtre d'origine
    driver.switch_to.window(driver.window_handles[0])

    # Cliquer sur "Lancer la recherche"
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, 'ctl0_CONTENU_PAGE_AdvancedSearch_lancerRecherche'))
    )
    lancer_recherche_button = driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_AdvancedSearch_lancerRecherche')
    lancer_recherche_button.click()
    # Attendre que l'élément select soit visible
    select_option = driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_resultSearch_listePageSizeBottom')
    
    # Utiliser Select pour manipuler l'élément select
    select = Select(select_option)
    
    # Sélectionner une option par sa valeur
    select.select_by_value('50')  # Sélectionner l'option avec la valeur "50"
    

    elements = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'img[alt="Accéder à la consultation"]'))
    )

    # Liste pour stocker les URLs des liens à traiter
    links_to_process = []

    # Collecter les URLs des liens dans la liste
    for element in elements:
        link = element.find_element(By.XPATH, './ancestor::a')
        links_to_process.append(link.get_attribute('href'))

    # Parcourir chaque lien pour effectuer le traitement
    for link in links_to_process:
        try:
            # Ouvrir le lien pour accéder à la page de consultation
            driver.get(link)

            # Attendre que le lien de téléchargement du dossier de consultation soit visible
            download_link = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.ID, 'ctl0_CONTENU_PAGE_linkDownloadDce'))
            )

            # Obtenir l'URL du lien de téléchargement
            download_url = download_link.get_attribute('href')

            # Afficher l'URL du dossier de consultation à télécharger
            print("URL du dossier de consultation à télécharger :", download_url)

            # Naviguer vers l'URL du lien de téléchargement
            driver.get(download_url)

            # Attendre que la case à cocher pour accepter les conditions soit visible et la cocher
            checkbox = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_accepterConditions'))
            )
            checkbox.click()

            # Remplir les champs du formulaire (exemples ici)
            nom_field = driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_nom')
            prenom_field = driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_prenom')
            email_field = driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_EntrepriseFormulaireDemande_email')

            nom_field.send_keys("Dupont")
            prenom_field.send_keys("Jean")
            email_field.send_keys("jean.dupont@example.com")

            # Cliquer sur le bouton "Valider"
            valider_button = driver.find_element(By.ID, 'ctl0_CONTENU_PAGE_validateButton')
            valider_button.click()

            # Attendre que le lien de téléchargement soit disponible et cliquer dessus
            download_dossier_link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'ctl0_CONTENU_PAGE_EntrepriseDownloadDce_completeDownload'))
            )
            download_dossier_link.click()

            # Attendre que le fichier soit téléchargé
            time.sleep(10)  # Peut-être ajuster le temps d'attente en fonction de la taille du fichier

            # Trouver le fichier zip téléchargé
            files = os.listdir(download_dir)
            zip_file = None
            for file in files:
                if file.endswith('.zip'):
                    zip_file = file
                    break

            if zip_file:
                zip_path = os.path.join(download_dir, zip_file)

            # Décompresser le fichier zip
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # Extraire les fichiers dans un dossier temporaire
                    zip_ref.extractall(download_dir)

            # Supprimer le fichier zip après extraction
                os.remove(zip_path)

            # Parcourir les fichiers extraits (y compris les fichiers dans les sous-dossiers)
                for root, dirs, files in os.walk(download_dir):
                    for file in files:
                        if any(file.endswith(extension) for extension in ['.pdf', '.doc', '.docx']) and re.match(r'^CPS', file, re.IGNORECASE):
                # Construire le chemin complet du fichier
                            file_path = os.path.join(root, file)
                # Déplacer le fichier PDF commençant par "CPS" vers le répertoire DATA
                            shutil.move(file_path, os.path.join(data_dir, file))

                print("Fichiers PDF commençant par 'CPS' déplacés vers le répertoire DATA avec succès.")

            else:
                print("Aucun fichier zip trouvé.")

            # Revenir à la page des résultats de recherche
            driver.back()
            driver.back()  # Revenir deux fois en arrière

                # Cliquer sur le lien pour retourner aux offres
            link_retour = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'ctl0_CONTENU_PAGE_linkRetourBas2'))
                )
            link_retour.click()



        except Exception as e:
            print(f"Une erreur s'est produite lors du traitement du lien {link} : {str(e)}")


finally:
    # Fermer le driver
    driver.quit()
