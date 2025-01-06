from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
import openpyxl
from time import sleep

def iniciar_driver():



    chrome_options = Options()

    arguments = ['--lang=pt-BR', '--window-size=800,600', '--incognito']

    for argument in arguments:
        chrome_options.add_argument(argument)

    caminho_padrao_download = 'C:\\Users\\Gabri\\Downloads'

    chrome_options.add_experimental_option("prefs", {
        'download.default_directory': caminho_padrao_download,

        'download.directory_upgrade': True,

        'download.prompt_for_download': False,

        "profile.default_content_setting_values.notifications": 2,

        "profile.default_content_setting_values.automatic_downloads": 1,
    })

    driver = webdriver.Chrome(options=chrome_options)
    return driver

# Definindo workbook(excel file):
workbook = openpyxl.load_workbook('arquivos\\produtos_ficticios.xlsx')
tabela_produtos = workbook['Produtos']

# Iniciando tela do site
driver = iniciar_driver()
driver.get('https://cadastro-produtos-devaprender.netlify.app')
driver.maximize_window()
sleep(5)

# Preenchimento:
for linha in tabela_produtos.iter_rows(min_row=2, max_row=5):
    # Product name
    name_field = driver.find_element(By.XPATH,"//input[@id='product_name']")
    name_field.send_keys(linha[0].value)
    sleep(1)

    # Description
    description_field = driver.find_element(By.XPATH,"//textarea[@id='description']")
    description_field.send_keys(linha[1].value)
    sleep(1)
    
    # Category
    category_field = driver.find_element(By.XPATH,"//input[@id='category']")
    category_field.send_keys(linha[2].value)
    sleep(1)

    # Product Code
    code_field = driver.find_element(By.XPATH,"//input[@id='product_code']")
    code_field.send_keys(linha[3].value)
    sleep(1)

    # Weight(Kg)
    weight_field = driver.find_element(By.XPATH,"//input[@id='weight']")
    weight_field.send_keys(linha[4].value)
    sleep(1)

    # Dimensions(Size)
    dimension_field = driver.find_element(By.XPATH,"//input[@id='dimensions']")
    dimension_field.send_keys(linha[5].value)
    sleep(1)

    # Click in next button
    chain = ActionChains(driver)
    chain.send_keys(Keys.TAB).pause(1).send_keys(Keys.ENTER).perform()
    sleep(2)

    # Price
    price_field = driver.find_element(By.XPATH,"//input[@id='price']")
    price_field.send_keys(linha[6].value)
    sleep(1)
     
    # Qtd Stock
    qtd_field = driver.find_element(By.XPATH,"//input[@id='stock']")
    qtd_field.send_keys(linha[7].value)
    sleep(1)

    # Expiry Date
    expiry_field = driver.find_element(By.XPATH,"//input[@id='expiry_date']")
    expiry_field.send_keys(linha[8].value)
    sleep(1)

    # Color
    color_field = driver.find_element(By.XPATH,"//input[@id='color']")
    color_field.send_keys(linha[9].value)
    sleep(1)

    # Size
    size_dropdown = driver.find_element(By.XPATH,"//select[@id='size']")
    size_dropdown.click()
    sleep(1)

    if (linha[10].value) == 'Pequeno':
        chain.send_keys(Keys.ENTER).perform()
    elif (linha[10].value) == 'Médio':
        chain.send_keys(Keys.DOWN).pause(0.5).send_keys(Keys.ENTER).perform()
    elif (linha[10].value) == 'Grande':
        chain.send_keys(Keys.DOWN).pause(0.5).send_keys(Keys.DOWN).pause(0.5).send_keys(Keys.ENTER).perform()
    sleep(1)

    # Material
    material_field = driver.find_element(By.XPATH,"//input[@id='material']")
    material_field.send_keys(linha[11].value)
    sleep(1)

    # Click next button
    chain.send_keys(Keys.TAB).pause(0.5).send_keys(Keys.ENTER).pause(4).perform()

    # Manufacturer
    manufacturer_field = driver.find_element(By.XPATH,"//input[@id='manufacturer']")
    manufacturer_field.send_keys(linha[12].value)
    sleep(1)

    # Origin country
    country_field = driver.find_element(By.XPATH,"//input[@id='country']")
    country_field.send_keys(linha[13].value)
    sleep(1)

    # Observations
    obs_field = driver.find_element(By.XPATH,"//textarea[@id='remarks']")
    obs_field.send_keys(linha[14].value)
    sleep(1)

    # Barcode
    barcode_field = driver.find_element(By.XPATH,"//input[@id='barcode']")
    barcode_field.send_keys(linha[15].value)
    sleep(1)

    # WareHouse Location
    location_field = driver.find_element(By.XPATH,"//input[@id='warehouse_location']")
    location_field.send_keys(linha[16].value)
    sleep(1)

    # Click in conclude button
    chain.send_keys(Keys.TAB).pause(0.5).send_keys(Keys.ENTER).perform()
    sleep(1)

    # Click in ok button alert
    alert1 = driver.switch_to.alert
    alert1.accept()
    sleep(2)

    # Click in "add one more"
    chain.send_keys(Keys.TAB).pause(0.5).send_keys(Keys.ENTER).pause(3).perform() 


input('')


