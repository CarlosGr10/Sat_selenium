from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import ElementNotInteractableException
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

    # segunda ventana (Addendas y Complementos (CFDI 4.0))
    driver.get("https://www.mysuitecfdi.com/Home_Operaciones40.aspx")

    # Tercera ventana (Complemento de pago 2.0 (REP))
    driver.get("https://www.mysuitecfdi.com/Facturar_Pagos40.aspx")

    # Tab de Emisor
    ddl_regimen_fiscal = driver.find_element(By.XPATH, "//*[@id=\"ddlRegimenFiscal\"]")
    select = Select(ddl_regimen_fiscal)

    # esperar hasta que el elemento esté visible
    driver.implicitly_wait(10) # espera hasta 10 segundos

    ddl_regimen_fiscal = driver.find_element(By.XPATH, "//*[@id=\"ddlRegimenFiscal\"]")
    select = Select(ddl_regimen_fiscal)

    select.select_by_value("601")
    
    # Tab de Receptor
    
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id=\"tabs\"]/ul/li[2]"))).click()
    rfc = driver.finrfc = driver.find_element("xpath", "//*[@id=\"txtR_RFC\"]")
    rfc.send_keys("CEC8504255S6")
    rfc.send_keys(Keys.ENTER)

    razon_social = driver.finrfc = driver.find_element("xpath","//*[@id=\"txtR_Name\"]")
    razon_social.send_keys("vale")
    razon_social.send_keys(Keys.ENTER)

    
    correo = driver.finrfc = driver.find_element("xpath","//*[@id=\"txtR_Mail\"]") 
    correo.send_keys("vale@algo.com")
    correo.send_keys(Keys.ENTER)
    
    domicilio_cp = driver.finrfc = driver.find_element("xpath","//*[@id=\"txtDomicilioFiscalReceptor\"]")
    domicilio_cp.send_keys("1123")
    domicilio_cp.send_keys(Keys.ENTER)

    driver.implicitly_wait(10) # espera hasta 10 segundos
    regimen_receptor = driver.find_element(By.XPATH, "//*[@id=\"ddlRegimenFiscalReceptor\"]")
    select = Select(regimen_receptor)

    select.select_by_value("601")


    # Tab Datos del complemento
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id=\"tabs\"]/ul/li[5]"))).click()

    time.sleep(1)

    # Adicionar pagos
    button_adiciona_pago = driver.find_element("xpath", "//*[@id=\"btnAddPago\"]")
    button_adiciona_pago.click()

    fecha_pago = driver.find_element(By.ID, "txtP_FechaPago")
    fecha_pago.send_keys("2023-03-16 12:00:00")
    fecha_pago.send_keys(Keys.ENTER)

    time.sleep(1)
    
    # Ciclo Impuestos P. Traslados
    driver.implicitly_wait(10)
    wait = WebDriverWait(driver, 10)
    button_impuesto_traslado = wait.until(EC.element_to_be_clickable((By.ID, 'btnAddTrasladoP')))
    attempts = 0
    while attempts < 5:
        try:
            button_impuesto_traslado.click()
            break
        except ElementNotInteractableException:
            attempts += 1

    driver.implicitly_wait(10)
    base = driver.finrfc = driver.find_element(By.XPATH, '//*[@id="TrasladoP_BaseP"]')
    base.send_keys(1)
    base.send_keys(Keys.ENTER)
   

    driver.implicitly_wait(10)
    impuesto_p = driver.find_element(By.XPATH, "//*[@id=\"TrasladoP_ImpuestoP\"]")
    select = Select(impuesto_p)
    select.select_by_value("002")

    driver.implicitly_wait(10)
    tipo_factor = driver.find_element(By.XPATH, "//*[@id=\"TrasladoP_TipoFactorP\"]")
    select = Select(tipo_factor)
    select.select_by_value("Tasa")

    # Boton Guardar (Ventana emergente)
    wait = WebDriverWait(driver, 10)
    element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".ui-dialog-buttonpane")))
    guardar_button = element.find_element(By.XPATH, "//button[text()='Guardar']")
    guardar_button.click()

    
    # Ciclo Documentos relacionados
    wait = WebDriverWait(driver, 10)
    button_adicionar_documentos = driver.find_element("xpath", "//*[@id=\"btnAddDoc\"]")
    button_adicionar_documentos.click()

    id_doc = driver.finrfc = driver.find_element("xpath",'//*[@id="txtPDoc_IdDocumento"]')
    id_doc.send_keys("HOLAAAAAA")
    id_doc.send_keys(Keys.ENTER)
    
    serie = driver.finrfc = driver.find_element("xpath",'//*[@id="txtPDoc_Serie"]')
    serie.send_keys("A")
    serie.send_keys(Keys.ENTER)
    
    folio = driver.finrfc = driver.find_element("xpath",'//*[@id="txtPDoc_Folio"]')
    folio.send_keys("B")
    folio.send_keys(Keys.ENTER)
    
    equivalencia = driver.finrfc = driver.find_element("xpath",'//*[@id="txtPDoc_TipoCambioDR"]')
    equivalencia.send_keys("1")
    equivalencia.send_keys(Keys.ENTER)
    
    num_parcialidad = driver.finrfc = driver.find_element("xpath",'//*[@id="txtPDoc_NumParcialidad"]')
    num_parcialidad.send_keys("2")
    num_parcialidad.send_keys(Keys.ENTER)
    
    sueldo_anterior = driver.finrfc = driver.find_element("xpath",'//*[@id="txtPDoc_ImpSaldoAnt"]')
    sueldo_anterior.send_keys("3")
    sueldo_anterior.send_keys(Keys.ENTER)
    
    imp_pagado = driver.finrfc = driver.find_element("xpath",'//*[@id="txtPDoc_ImpPagado"]')
    imp_pagado.send_keys("4")
    imp_pagado.send_keys(Keys.ENTER)
    
    sueldo_insoluto = driver.finrfc = driver.find_element("xpath",'//*[@id="txtPDoc_ImpSaldoInsoluto"]')
    sueldo_insoluto.send_keys("5")
    sueldo_insoluto.send_keys(Keys.ENTER)

    driver.implicitly_wait(10)
    moneda = driver.find_element(By.XPATH, '//*[@id="ddlPDoc_MonedaDR"]')
    select = Select(moneda)
    select.select_by_value("MXN|2")

    driver.implicitly_wait(10)
    obj_im_dr = driver.find_element(By.XPATH, '//*[@id="ddlPDoc_ObjetoImpDR"]')
    select = Select(obj_im_dr)
    select.select_by_value("03")

    buttom_adicion = guardar_button = element.find_element("xpath", '//*[@id="btnAddOkDoc"]')
    buttom_adicion.click()

    
    
    time.sleep(2)


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

