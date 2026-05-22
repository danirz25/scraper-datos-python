import pandas as pd
import requests
import os
from io import  StringIO


url = input("Ingresa el link de la página donde deseas descargar las tablas: ")

#Hya paginas que rechazan la solicitud pq reconcoen que es un script, entonces se disfraza la petición:
# Usamos un "User-Agent". Esto es el "disfraz" para parecer un navegador real.
header = {
  "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
}

try:
    # Hacemos la petición a la web 
    r = requests.get(url, headers=header)
    
    # Le pasamos el TEXTO del html a pandas (en vez de la url directa, no se pq pero con url dejó de funcionar xd)
    tablas = pd.read_html(r.text)

    # Se buscan todas las tablas
    #tablas = pd.read_html(url)
    #Cuando encuentre una tabla, 
    if len(tablas) > 0:
        while True:
          #Se le pide el nombre del archivo al usuario para que no se sobreescriba nada 
          nombre_usuario =input("Elige el nombre del archivo: ").strip()  

            #Acá se verifica que el usuario no ponga un punto
          if not nombre_usuario or nombre_usuario == ".":
            print("ERROR: No se puede colocar un caracter de ese tipo. Intenta nuevamente. ")
            continue
          
           #Acá se verifica si se añadió la extensión .xlsx, si no (lo más logico), se añade
          if not nombre_usuario.endswith(".xlsx"):
              nombre_archivo= nombre_usuario + ".xlsx"
          else:
              nombre_archivo = nombre_usuario
          
            #Verificación de que el nombre que se haya puesto no esté repetido.
          if os.path.exists(nombre_archivo):
              print("ERRROR: Un archivo con este nombre ya existe. Por favro, intente nuevamente. ")
              continue
          break

        #nombre_archivo = "Prueba.xlsx"
        
        #exclewriter es para que casa tabla sea una hoja (porque en la version pasada solo te considera la ultima)
        with pd.ExcelWriter(nombre_archivo) as writer:
            
            #Ciclo for para iterar en cada tabla 
            for i, tabla in enumerate(tablas):
                # Creamos un nombre 
                nombre_hoja = f"Tabla_{i+1}"
                
                # Guardamos cada tabla en su hoja correspondiente
                tabla.to_excel(writer, sheet_name=nombre_hoja, index=False)
                
        print(f"Se descargaron {len(tablas)} tablas y se guardaron en '{nombre_archivo}'.")

    else:
        print("No se encontraron tablas en la página.")
except Exception as e:
    print(f"Ocurrió un error al intentar acceder a la página: {e}")

##El problema de este código es que solo copia las tablas estáticas, las dinámicas no las considera.