import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import pandas as pd

#config do driver
download_dir = os.path.expanduser("~/Downloads") # diretorio de downloads do pc
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": r"C:\Users\PC\Downloads",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": False,
    "profile.default_content_settings.popups": 0
})

pyautogui.PAUSE = 0.5


#inicializar driver
driver = webdriver.Chrome()
driver.get("https://08978175000180.portal-veri.com.br")
driver.set_window_size(1366, 768)


#LOGIN NO DRIVE
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_sign_in_form"]/div[2]/input'))).send_keys("joas@controllersbr.com") # Ajuste conforme necessário
time.sleep(0.5)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_sign_in_form"]/div[3]/input'))).send_keys("Jd@005570") # Ajuste conforme necessário
time.sleep(0.5)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_sign_in_form"]/div[5]/button'))).click()

#clica no elemento dash federal
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_app_content_container"]/div/div/div[2]/div[1]/div[3]/a/div/span'))).click()
time.sleep(0.5)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-id_empresa_1-container"]/span'))).click()
time.sleep(1)

tabela = pd.read_csv(r"C:\Users\PC\Desktop\Captura de relatório por cliente sistema veri\CONSULTAR NO VERI.csv")
for linha in tabela.index:
    empresa = tabela.loc[linha, 'cnpj']
    pyautogui.write(empresa)
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_app_content_container"]/div/div/div[2]/button'))).click()
    time.sleep(0.5)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_modal_export_users"]/div/div/div[2]/div[1]/span/span[1]/span'))).click()
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.press('down')
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.press('esc')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-id_empresa_1-container"]/span'))).click()
    time.sleep(0.5)
    


   
    
