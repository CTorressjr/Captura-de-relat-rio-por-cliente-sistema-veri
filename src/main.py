import os
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
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

#parametros
pyautogui.PAUSE = 0.5
driver = webdriver.Chrome()
tabela = pd.read_csv(r"C:\Users\torre\Desktop\Captura-de-relat-rio-por-cliente-sistema-veri\cnpj.csv")

#inicializar driver
def initialize():
    driver.get("https://CNPJ.portal-veri.com.br")
    driver.set_window_size(1366, 768)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_sign_in_form"]/div[2]/input'))).send_keys("LOGIN") # Ajuste conforme necessário
    time.sleep(0.5)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_sign_in_form"]/div[3]/input'))).send_keys("SENHA") # Ajuste conforme necessário
    time.sleep(0.5)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_sign_in_form"]/div[5]/button'))).click()
    selectdash()
#clica no elemento dash federal
def selectdash():
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_app_content_container"]/div/div/div[2]/div[1]/div[3]/a/div/span'))).click()
    time.sleep(0.5)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="select2-id_empresa_1-container"]/span'))).click()
    time.sleep(1)
    capturarelatorio()


def capturarelatorio():
    for linha in tabela.index:
        empresa = tabela.loc[linha, 'cnpj']
        pyautogui.write(str(empresa))
        pyautogui.press('enter')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_app_content_container"]/div/div/div[2]/button'))).click()
        time.sleep(0.5)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_modal_export_users"]/div/div/div[2]/div[1]/span/span[1]/span'))).click()
        time.sleep(1)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_modal_export_users"]/div/div/div[2]/div[1]/span/span[1]/span'))).click()
        time.sleep(3)
        actions = ActionChains(driver)
        actions.key_down(Keys.ARROW_DOWN).perform()    
        time.sleep(1)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_modal_export_users"]/div/div/div[2]/div[2]/button/span[1]'))).click()
        time.sleep(0.5)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="kt_modal_export_users"]/div/div/div[1]/div/i'))).click()
        time.sleep(0.5)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_app_header_wrapper"]/span/span[1]/span'))).click()
        time.sleep(0.5)
    

initialize()
   


