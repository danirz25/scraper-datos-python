#Esta versión de Scraping contiene una iterfaz gráfica para que sea más cómoda de usar para cualquiera
import pandas as pd
import os
import tkinter as tk
from tkinter import messagebox, filedialog
#import time 
from io import  StringIO
from selenium import webdriver      #Selenium es para poder trabajar con JavaScripts en la web
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options


class ScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Descargador de Tablas Web")
        self.root.geometry("500x350")
        self.driver = None

        self.lbl_instruction = tk.Label(root, text = "Ingresa el link para descargar sus tablas.")                                                     #instrucciones 
        self.lbl_instruction.pack(pady=(20,5))

        self.entry_url = tk.Entry(root, width=50)                                                                                                      #URL caja de texto
        self.entry_url.pack(pady=5)

        self.btn_abrir = tk.Button(root, text = "Buscar y verificar tablas", command=self.abrir_navegador, bg="#4CAF50", fg="white", height=2)         #Boton para abrir el link
        self.btn_abrir.pack(pady=10)

        self.lbl_estado = tk.Label(root, text="Esperando URL...", fg="gray", wraplength=450)                                                           #Etiqueta de estado esperando 
        self.lbl_estado.pack(pady=10)
        
        self.btn_descargar = tk.Button(root, text="Verificado. Descargar Excel", command=self.descargar_tablas, state="disabled", bg="#008CBA", fg="white", height=2)
        self.btn_descargar.pack(pady=20)                                                                                                               #Botón para descargar
    
    
    def abrir_navegador(self):                                                                  #FUNCION Para abrir el chrome                                 
        url = self.entry_url.get().strip()

        if not url:
            messagebox.showwarning("Error", "Por favor ingresa una URL válida.")                 #Por si el link está mal
            return

        try:
            #Para las tablas dinámicas, se cambia requests por selenium. 
            chrome_options = Options()
            # chrome_options.add_argument("--headless")   #Se quita para que se abra el navegador y poder cliquear 
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")

            self.lbl_estado.config(text="Iniciando... Por favor espera.", fg="blue")             # Iniciamos el driver
            self.root.update()                                                                   # Forzar actualización de la interfaz visual

            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)     #Se incia el navegador
            self.driver.get(url)                                                                                         #Se obtiene el link

            # Se actualiza la interfaz para poder descargar después 
            self.lbl_estado.config(text="Navega en la ventana hasta tener las tablas que desees.\nCuando las tengas, presiona el botón de abajo.", fg="black")
            self.btn_descargar.config(state="normal")                                                              #habilitamos el segundo botón
            self.btn_abrir.config(state="disabled")                                                                #deshabilitamos el primero para evitar problemas

        except Exception as e:                                                                                     #Excepción por si hay error
            messagebox.showerror("Error", f"No se pudo abrir el navegador:\n{e}")
            self.lbl_estado.config(text="Error al abrir navegador.")

    def descargar_tablas(self):                                                                #Función para descargar las tablas
        if not self.driver:
            return

        try:
            html_completo = self.driver.page_source   #Obtenemos el html completo
            
            self.driver.quit()                        #Se cierra el navegador 
            self.driver = None
            
            html_content = StringIO(html_completo)    # Procesamos con Pandas
            tablas = pd.read_html(html_content)       # Le pasamos el TEXTO del html a pandas (en vez de la url directa, no sé por qué pero con url dejó de funcionar xd)

            if len(tablas) == 0:
                messagebox.showwarning("Aviso", "No se encontraron tablas en la página actual.")   #mensaje por si no encuentra tablas 
                self.reiniciar_interfaz()
                return

            archivo_guardar = filedialog.asksaveasfilename(         #Ahora no pido el nombre, directamente se abre el explorador de archivos 
                defaultextension=".xlsx",           
                filetypes=[("Archivos de Excel", "*.xlsx"),
                           ("Archivos CSV", "*.csv" )
                          ],                                        #Igual se pone por defecto la extensión y tipo de archivo
                title="Guardar tablas como..."                      #Añadir / verificar que se pueda cambiar la extensión
            )

            if not archivo_guardar:                                 #Si el usuario cancela la ventana de guardar
                messagebox.showinfo("Cancelado", "No se guardó el archivo.")  
                self.reiniciar_interfaz()
                return

            guardados = 0
            mensaje_final = ""

            if archivo_guardar.endswith(".xlsx"):
                with pd.ExcelWriter(archivo_guardar) as writer:         #exclewriter es para que cada tabla descargada sea una hoja de excel 
                    for i, tabla in enumerate(tablas):                  #Ciclo for para iterar en cada tabla
                     nombre_hoja = f"Tabla_{i+1}"                    #A cada tabla se le da el nombre 'Tabla i'
                    
                     if isinstance(tabla.columns, pd.MultiIndex):    #se añade esto para 'aplanar' las celdas combinada porque si no, da error.
                        tabla.columns = [' '.join(map(str, col)).strip() for col in tabla.columns.values]
                    
                     if not tabla.empty:                             #El if es para evitar guardar tablas vacias 
                        tabla.to_excel(writer, sheet_name=nombre_hoja, index=False)   # Guardamos cada tabla en su hoja correspondiente
                        guardados += 1
                messagebox.showinfo("Éxito",f"Se descargaron {len(tablas)} tablas correctamente en: \n{archivo_guardar}")

            elif archivo_guardar.endswith(".csv"):
                base_path = os.path.splitext(archivo_guardar)[0]
                
                for i, tabla in enumerate(tablas):
                    if isinstance(tabla.columns, pd.MultiIndex):    #se añade esto para 'aplanar' las celdas combinada porque si no, da error.
                        tabla.columns = [' '.join(map(str, col)).strip() for col in tabla.columns.values]
                    if not tabla.empty:
                        if len(tablas) > 1:
                            nombre_csv = f"{base_path}_Tabla_{i+1}.csv"
                        else:
                            nombre_csv = archivo_guardar
                        
                        tabla.to_csv(nombre_csv, index=False, encoding='utf-8-sig')  #UTF 8 Es para que reconozca acentos y la ñ
                        guardados += 1
            if len(tablas) >1:
                mensaje_final = f"Al ser formato CSV, se genereron {guardados} archivos separados (Uno por tabla)."
            else:
                mensaje_final = f"Se guardó correctamente el archivo CSV."

            messagebox.showinfo("Éxito", f"{mensaje_final}\nUbicación: {os.path.dirname(archivo_guardar)}")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al procesar los datos:\n{e}")
        
        finally:
            self.reiniciar_interfaz()

    def reiniciar_interfaz(self):                                   #Esto es para que no se quede con los botones bloqueados y siga funcionando 
        """Devuelve los botones a su estado original para hacer otra descarga"""
        self.btn_abrir.config(state="normal")
        self.btn_descargar.config(state="disabled")
        self.lbl_estado.config(text="Listo para nueva descarga.", fg="green")
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None

#Bloque de ejecución
if __name__ == "__main__":
    ventana = tk.Tk()
    app = ScraperApp(ventana)
    ventana.mainloop()


#Para Comentar varias lienas es Ctrl + K + C 
#Para Descomentar varias lineas es Ctrl + K + U 

