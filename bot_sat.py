from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import os
import time



def page_web():
    # Este es el webdriver definido
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    # Link de la pagiina web
    driver.get("https://www.mysuitecfdi.com/login.aspx")

     # Esta es la primera ventana
    rfc = driver.find_element("xpath", "//*[@id=\"inputRFC\"]")
    rfc.send_keys("KDS8907241B8")
    rfc.send_keys(Keys.ENTER)

    user = driver.finrfc = driver.find_element("xpath", "//*[@id=\"inputUser\"]")
    user.send_keys("Cobranza_2")
    user.send_keys(Keys.ENTER)

    password = driver.finrfc = driver.find_element("xpath", "//*[@id=\"Login100_Password\"]")
    password.send_keys("Camilin8#")
    password.send_keys(Keys.ENTER)

    # espera hasta que el elemento sea visible
    boton_login = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, '//*[@id="Login100_LoginButton"]'))
    )

    # hace clic en el elemento
    boton_login.click()

    # segunda ventana

    addendas_element = driver.find_element_by_id('ContentPlaceHolder1_Addendas')
    addendas_url = addendas_element.get_attribute('href')
    addendas_url.click()


    '''
    se le da click ===> Addendas y Complementos (CFDI 4.0)
    se le da click a ===> Complemento de pago 2.0 (REP)
    Se le da click al tab de Receptor ==> Llenar datos
    se salta a al tab Datos del complemento ===> LLenar datos

    Realizar un cliclo en (Impuestos P. Traslados)[Ciclo direto]

    Llenar Documentos Relacionados
    Dar click en ===> Completar adicion/adicion del doc [ciclo anidado]

    Generar CDFI



    '''

if __name__ == "__main__":
    page_web()

