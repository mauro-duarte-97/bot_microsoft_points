import time
import ctypes
import psutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from random_word import RandomWords
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service

#CONSTANTE ALERTA
MB_SYSTEMMODAL = 0x00001000

#FUNCIONES

def generador_lista_palabras():
    r = RandomWords()
    # Generar una palabra al azar
    lista_temas = []
    for i in range(0, 30):
        palabra_azar = r.get_random_word()
        lista_temas.append(palabra_azar)
    return lista_temas

def openWeb(driver, listaTematicas):

    try:
        driver.get('https://www.bing.com/search?q=Bing+AI&qs=HS&sc=11-0&cvid=AD62DD36244A4856B319047D747E7840&FORM=QBLH&sp=1&lq=0')
        driver.implicitly_wait(5)
        time.sleep(5)
        
        #Inserto dato en el buscador
        buscador = driver.find_element(By.XPATH, '/html/body/header/form/div/textarea') 
        try:
            buscador.clear()
        except:
            pass
        buscador.send_keys('Noticiero')
        buscador.submit()
        time.sleep(5)

        # Ciclo de busqueda
        for i in range(0, 30):
            # Busca una tematica de una lista de tematicas
            try:
                buscador = driver.find_element(By.XPATH, '/html/body/header/form/div/textarea') # Tengo que declararlo de nuevo porque el selector se vence o se vuelve inestable
                buscador.clear()
            except:
                pass
            buscador.send_keys(listaTematicas[i])
            buscador.submit()
            time.sleep(10)


    except:
	    ctypes.windll.user32.MessageBoxW(0, "El bot fallo al ejecutar", "BOT ERROR", MB_SYSTEMMODAL)
         

def kill_edgedriver():
    # Cierra todos los procesos de msedgedriver y Microsoft Edge
    try:
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] in ["msedgedriver.exe", "msedge.exe"]:
                proc.kill()
    except:
        print("Error al cerrar los procesos de msedgedriver y Microsoft Edge")


def driver_init():
    # Ruta del perfil de usuario de Edge
    user_data_dir = r'C:/Users/Mauro/AppData/Local/Microsoft/Edge/User Data'

    # Ruta al driver descargado
    driver_path = r'D:/Proyectos/python/chromedriver/msedgedriver.exe'

    # Crear un objeto de opciones para Edge
    edge_options = webdriver.EdgeOptions()

    # Configurar el uso del perfil de usuario existente
    edge_options.add_argument(f'--user-data-dir={user_data_dir}')
    edge_options.add_argument('--profile-directory=Default')  # Cambia "Default" si usas otro perfil

    # Maximiza la ventana
    edge_options.add_argument('--start-maximized')

    # Configurar las opciones como desees (este es solo un ejemplo)
    # edge_options.use_chromium = True

    # Crear servicio con el driver
    service = Service(executable_path=driver_path)

    kill_edgedriver()

    # Crear una nueva instancia del driver de Edge con las opciones configuradas
   # Inicializar el navegador
    driver = webdriver.Edge(service=service, options=edge_options)
    return driver


def main():

    lista_temas = generador_lista_palabras()

    # Carga el driver
    driver = driver_init()

    openWeb(driver, lista_temas)

    driver.quit()


#### EJECUCION ####
main()