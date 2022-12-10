from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas
import autoit
import easygui as eg





def abrir_whatsapp()->webdriver:    
    options = webdriver.ChromeOptions() 
    options.add_argument("user-data-dir=C:\\Users\\sQreen\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 1")    
    driver = webdriver.Chrome(executable_path='webdriver//chromedriver.exe',chrome_options=options)
    driver.get('https://web.whatsapp.com/')
    driver.maximize_window()
    eg.msgbox("Ingrese el código QR, y espere hasta que cargue la aplicación para continuar.","QR","Ok")    
    return driver
     

def send_img(driver,phone:str,img:str,sms:str)->None:
    try:
        activar_busqueda = driver.find_element(By.XPATH,'//*[@id="side"]/div[1]/div/div/button/div[2]/span').click()  
        time.sleep(2)        
        busqueda = driver.find_element(By.XPATH,'//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]')
        time.sleep(2)
        busqueda.send_keys(phone)    
        time.sleep(2)
        busqueda.send_keys(Keys.ENTER)
        time.sleep(3)
        adjuntar = driver.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/span').click()
        time.sleep(2)
        adjuntar_imagen = driver.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[1]/button/span')
        time.sleep(2)
        adjuntar_imagen.click()
        time.sleep(2)
        autoit.win_wait("Abrir",30)
        time.sleep(2)
        autoit.win_active("Abrir")
        time.sleep(1)
        autoit.control_set_text("Abrir","Edit1",img)
        time.sleep(1)
        autoit.control_send("Abrir","Edit1","{ENTER}")
        time.sleep(2)
        pie_foto = driver.find_element(By.XPATH,'//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[1]/div[3]/div/div/div[2]/div[1]/div[1]/p').send_keys(sms)    
        time.sleep(3)    
        #btn_enviar = driver.find_element(By.XPATH,'//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div').click()
        driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div/span').click()
        time.sleep(5)
    except Exception as e:
        print(e)
        time.sleep(5)
    #eg.msgbox("Pausa por red lenta.","Pausa de prueba","Ok")


def send_text(driver,phone:str,texto:str)->None:
    activar_busqueda = driver.find_element(By.XPATH,'//*[@id="side"]/div[1]/div/div/button/div[2]/span').click()  
    time.sleep(2)        
    busqueda = driver.find_element(By.XPATH,'//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]')
    time.sleep(2)
    busqueda.send_keys(phone)    
    time.sleep(2)
    busqueda.send_keys(Keys.ENTER)
    time.sleep(3)
    parrafo = driver.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]').click() 
    time.sleep(2)   
    driver.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(texto)    
    time.sleep(2)
    btn_enviar = driver.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
    time.sleep(5)


def run()->None:
    driver = abrir_whatsapp()
    #df = pandas.read_excel('enviar_sms_whatsapp_grupos_datos.xlsx')
    ruta = "E:\\Soft\\python_venv\\Practicando_Python\\Ejercicios\\enviar_sms_whatsapp_grupos\\enviar_sms_whatsapp_grupos_dato.xlsx"
    df = pandas.read_excel(ruta)

    #txt = f"*COMPRAMOS ORO TODOS LOS KILATES*\t```8kt```\t```10kt```\t```14kt```\t```18kt```\t_Cadenas 10kt de fábrica entre 1g y 2.9g = 5000cup el gramo_\t_Cadenas 10kt de fábrica entre 3g y 4.9g = 4 375cup el gramo_\t_Cadenas 10kt 3 entre 5g y 9.9g = 4 125cup el gramo_\t*WhatsApp:* https://cutt.ly/UMKUkQC\t```-SI NECESITA HACER ESA INVERSIÓN...```\t```-COMPRAR ESA OTRA COSA QUE TANTO QUIERE ...```\t```-HACER ALGÚN VIAJE INTERNACIONAL ... ```\t```-CUMPLIR CON SU RELIGIÓN ...```\t```-DARSE ESAS VACACIONES QUE TANTO SE MERECE ```\t_Aproveche todo ese oro que día tras día a incrementado su valor. TENEMOS SU SOLUCIÓN.  No lo dude más contáctenos para la venta_ \t       *DE CLICK EN EL SIGUIENTE LINK* _WhatsApp:_ https://cutt.ly/UMKUkQC \t=>Anuncio Realizado por:\t*DPromo*, _Promocianamos tu anuncio en las Redes._ ```+53 53152384.```*¡Contáctenos!*"
    
    for index,celda in df.iterrows():
        try:
            if pandas.isna(celda['Ruta_Imagen']):
                print(f"{index} Insertando Texto {celda['Texto']} en {celda['Contacto']}")        
                send_text(driver=driver,phone=celda['Contacto'],texto=celda['Texto'])
            else:
                print(f"{index} Insertando Imagen {celda['Ruta_Imagen']} en {celda['Contacto']}")        
                send_img(driver=driver,phone=celda['Contacto'],img=celda['Ruta_Imagen'],sms=celda['Texto'])
        except Exception as e:
            print(e)
            pass

if __name__ == '__main__':
    run()