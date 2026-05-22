import pandas as pd
import os
import time 
from io import  StringIO
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options


#Para las tablas dinámicas, se cambia requests por selenium. 
chrome_options = Options()
# chrome_options.add_argument("--headless")   #Se quita para que se abra el navegador y poder cliquear 
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

url = input("Ingresa el link de la página donde deseas descargar las tablas: ") #Se pide el link para buscar y desbloquear tablas 
print("Iniciando Navegador. ")
print("Se abrirá una ventana")
print("Ve a la ventana que se abrió y busca la(s) tabla(s) que desees descargar.")
print("Una vez localizadas, vuelve a la terminal.")

try:
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)     #Se incia el navegador
    driver.get(url)                                                                                         #Se obtiene el link

    #En este bloque, tomamos el control para elegir la tabla en específico que queremos
    print("\n" + "="*60)
    input("Presiona [ENTER] AQUÍ en la terminal para descargar la(s) tabla(s).")
    print("="*60 + "\n")

    html_completo = driver.page_source   #Obtenemos el html completo
    driver.quit()                        #Se cierra el navegador 

    html_content = StringIO(html_completo)
    tablas = pd.read_html(html_content)     # Le pasamos el TEXTO del html a pandas (en vez de la url directa, no sé por qué pero con url dejó de funcionar xd)

    if len(tablas) > 0:
        while True:
          nombre_usuario =input("Elige el nombre del archivo: ").strip()  #Se le pide el nombre del archivo al usuario para que no se sobreescriba nada 

          if not nombre_usuario or nombre_usuario == ".":                 #Acá se verifica que el usuario no ponga un punto
            print("ERROR: No se puede colocar un caracter de ese tipo. Intenta nuevamente. ")
            continue
          
          if not nombre_usuario.endswith(".xlsx"):                        #Acá se verifica si se añadió la extensión .xlsx, si no (lo más logico), se añade
              nombre_archivo= nombre_usuario + ".xlsx"
          else:
              nombre_archivo = nombre_usuario
          
          if os.path.exists(nombre_archivo):                              #Verificación de que el nombre que se haya puesto no esté repetido.
              print("ERRROR: Un archivo con este nombre ya existe. Por favro, intente nuevamente. ")
              continue
          break
        
        #exclewriter es para que casa tabla sea una hoja (porque en la version pasada solo te considera la ultima)
        with pd.ExcelWriter(nombre_archivo) as writer:

            #Ciclo for para iterar en cada tabla 
            for i, tabla in enumerate(tablas):
                nombre_hoja = f"Tabla_{i+1}"                                 # Creamos un nombre para cada tabla
                
                if isinstance(tabla.columns, pd.MultiIndex):                 #se añade esto para 'aplanar' las celdas combinada porque si no, da error.
                    tabla.columns=[' '.join(map(str, col)).strip() for col in tabla.columns.values]
               
                tabla.to_excel(writer, sheet_name=nombre_hoja, index=False)  # Guardamos cada tabla en su hoja correspondiente
                
        print(f"Se descargaron {len(tablas)} tablas y se guardaron en '{nombre_archivo}'.")

    else:
        print("No se encontraron tablas en la página.")
except Exception as e:
    if 'driver' in locals():     #Cerramos el driver por cualquier error 
        driver.quit()
    print(f"Ocurrió un error al intentar acceder a la página: {e}")