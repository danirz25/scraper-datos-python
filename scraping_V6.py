#Esta versión deja de lado un poco la búsqueda en PDF(Necesita modificaciones severas)
#Acá se trata de descargar csv´s de forma automática datasets y pasarlas a excel o csv.
#Igualmente se da una mejora a la interfaz 
import pandas as pd
import requests
import os
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog
import pdfplumber                   #Para trabajar con PDF
from io import  StringIO            #Para tratar cadenas de texto como archivos
from selenium import webdriver      #Selenium es para poder trabajar con JavaScripts en la web
from selenium.webdriver.common.by import By #Para hacer búsquedas 
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

class ScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Descargador de Tablas Web - PDF")
        self.root.geometry("700x600")
        self.driver = None

        self.check_vars  = []
        self.enlaces_encontrados = []

        #Separación de funciones por pestañas
        self.tab_control = ttk.Notebook(root)            #Para que se distribuya mejor la infterfaz, se va a dividir en pestañas

        self.tab_web = tk.Frame(self.tab_control)        #PESTAÑA 1: Para sitios web
        self.tab_control.add(self.tab_web, text= ' Tablas en sitios Web ')

        self.tab_pdf = tk.Frame(self.tab_control)        #PESTAÑA 2: Para archivos PDF de forma local
        self.tab_control.add(self.tab_pdf, text= ' Archivos PDF ')
        
        self.tab_files = tk.Frame(self.tab_control)      #PESTAÑA 3: Para descargar archivos
        self.tab_control.add(self.tab_files, text= ' Archivos encontrados ')

        self.tab_control.pack(expand=1, fill="both")

        # 1 - PESTAÑA DE TABLAS EN SITIOS WEB 
        tk.Label(self.tab_web, text = "Ingresa el link para descargar sus tablas.", font =("Arial", 12, "bold")).pack(pady=(15,5))                     #instrucciones
        self.entry_url_web = tk.Entry(self.tab_web, width=50)                                                                                                      #URL caja de texto
        self.entry_url_web.pack(pady=5)
        #MODO: "WEB"                                                                                       
        tk.Button(self.tab_web, text = "Buscar y verificar tablas", command=lambda: self.abrir_navegador("WEB"), bg="#4CAF50", fg="white").pack(pady=(5))   #Boton para abrir el link 
        self.lbl_estado_web = tk.Label(root, text="Esperando URL...", fg="gray", wraplength=450)                                                           #Etiqueta de estado esperando 
        self.lbl_estado_web.pack(pady=5)
        self.btn_descargar_web = tk.Button(root, text="Verificado. Descargar Excel", command=self.descargar_tablas, state="disabled", bg="#008CBA", fg="white", height=2)
        self.btn_descargar_web.pack(pady=10)       #Botón para descargar

        # 2 - PESTAÑA DE TABLAS EN PDF's
        tk.Label(self.tab_pdf, text = "Selecciona el archivo PDF desde tu computadora: ", font=("Arial", 12, "bold")).pack(pady=(15,5))
        tk.Button(self.tab_pdf, text="Seleccionar", command = self.procesar_pdf, bg="#FF9800", fg="white", height=2).pack(pady=10)
        self.lbl_estado_pdf = tk.Label(self.tab_pdf, text="No se eligió ningún archivo", fg="gray")
        self.lbl_estado_pdf.pack(pady = 20)

        # 3 - PESTAÑA PARA BUSCAR Y DESCARGAR ARCHIVOS
        tk.Label(self.tab_files, text= " Descargar archivos CSV, XLSX Y ZIP ", font = ("Arial", 12, "bold")).pack(pady=(15,5))
        self.entry_url_files = tk.Entry(self.tab_files, width=60)
        self.entry_url_files.pack(pady=5) 
        tk.Button(self.tab_files, text=" Buscar archivos disponibles", command= lambda: self.abrir_navegador("ARCHIVOS"), bg = "#673AB7", fg="white").pack(pady=5)     #MODO: "ARCHIVOS"

        self.lbl_estado_files = tk.Label(self.tab_files, text= " Esperando URL . . .", fg="gray")
        self.lbl_estado_files.pack(pady=5)

        #Por si hay muchos archivos, se hace una lista con checkboxes para elegir los que se quieran descargar.
        self.frame_lista = tk.Frame(self.tab_files, bd = 1, relief="sunken")               #Parte gráfica con Tk
        self.frame_lista.pack(fill="both", expand=True, padx= 20, pady = 10)
        
        self.canvas = tk.Canvas(self.frame_lista, bg="white")
        self.scrollbar = tk.Scrollbar(self.frame_lista, orient= "vertical", command=self.canvas.yview)
        self.scroll_frame = tk.Frame(self.canvas, bg = "white")

        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0,0), window=self.scroll_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side = "left", fill="both", expand = True)
        self.scrollbar.pack(side="right", fill = "y")

        self.btn_download_files = tk.Button(self.tab_files, text ="Descargar archivos.", command=self.descargar_archivos_seleccionados, state= "disabled", bg="#008CBA", fg="white")
        self.btn_download_files.pack(pady = 10)
    
    #En esta función se implementa el modo, por si es busqueda WEB o descarga de ARCHIVOS
    def abrir_navegador(self, modo):                                                                  #FUNCION Para abrir el chrome                                 
        #url = self.entry_url.get().strip()                                     Ya no se usa, ahora se compara el modo.
        if modo == "WEB":
            url = self.entry_url_web.get().strip()
            lbl = self.lbl_estado_web
        else: 
            url = self.entry_url_files.get().strip()
            lbl= self.lbl_estado_files
        if not url:
            messagebox.showwarning("Error", "URL inválida.")                 #Por si el link está mal
            return

        try:
            #Para las tablas dinámicas, se cambia requests por selenium. 
            chrome_options = Options()
            chrome_options.add_argument("--start-maximized")
            #chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")

            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])  #Esto es algo anti bots para intentar entrar a links más protegidos
            chrome_options.add_experimental_option('useAutomationExtension', False)
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")

            lbl.config(text="Iniciando... Por favor espera.", fg="black")             # Iniciamos el driver
            self.root.update()                                                                   # Forzar actualización de la interfaz visual

            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)     #Se incia el navegador
            self.driver.get(url)                                                                                   #Se obtiene el link

            if modo == "WEB":
                lbl.config(text="Navega en la ventana hasta tener las tablas que desees.\nCuando las tengas, presiona el botón de abajo.", fg="black")
                self.btn_descargar_web.config(state="normal") 
            
            elif modo== "ARCHIVOS":
                lbl.config(text="Buscando archivos en la página. . .", fg="black")
                self.root.update()
                self.escanear_archivos()                                                       #Función para escanear archivos

        except Exception as e:                                                                 #Excepción por si hay error
            messagebox.showerror("Error", f"No se pudo abrir el navegador:\n{e}")
            if self.driver: self.driver.quit()

    def descargar_tablas(self):                                                                #Función para descargar las tablas
        if not self.driver:
            return

        try:
            html = self.driver.page_source            #Obtenemos el html completo
            self.driver.quit()                        #Se cierra el navegador 
            self.driver = None                        #Se optimizó esta ejec de lineas 97-101
            
            tablas = pd.read_html(StringIO(html))       # Le pasamos el TEXTO del html a pandas (en vez de la url directa, no sé por qué pero con url dejó de funcionar xd)

            if not tablas:                              #En lugar de comparar el num. de tablas ahora se hace directo xd 
                messagebox.showwarning("Aviso", "No se encontraron tablas en la página actual.")   #mensaje por si no encuentra tablas 
                return

            self.guardar_excel_csv(tablas, "WEB")       #Se hizo una función mejor para guardar tanto en web como html
            self.lbl_estado_web.config(text="Listo", fg="green")
            self.btn_descargar_web.config(state="disabled")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al procesar los datos:\n{e}")
        
        finally:
            self.reiniciar_interfaz()

    def procesar_pdf(self):                 #Función para procesar archivos PDF
        archivo = filedialog.askopenfilename (title = "Seleccionar Archivo PDF.", filetypes=[("PDF", "*.pdf")])
        if not archivo:
            return

        self.lbl_estado_pdf.config(text=f"Procesando: {os.path.basename(archivo)}...", fg="blue")
        self.root.update()
    
        try:
            tablas_totales = []
            with pdfplumber.open(archivo) as pdf:       #Se abre el pdf
                for pagina in pdf.pages:
                    extracto = pagina.extract_tables()   #Extrae tablas de la página actual

                    for t in extracto:
                        df = pd.DataFrame(t[1:], columns=t[0]).dropna(how='all')     #Convertir a Dataframe 
                        #df = df.dropna(how= 'all')                                  #Se asume que la primer fila son los encabezados
                    
                        if not df.empty:
                            tablas_totales.append(df)                           #Elimina filas totalmente vacías

            if tablas_totales:
                self.guardar_excel_csv(tablas_totales, "PDF")
                self.lbl_estado_pdf.config(text="Proceso realizado con éxito.", fg="green")
            else: 
                messagebox.showwarning("PDF","No se encontraron tablas en el archivo")
                self.lbl_estado_pdf.config(text="Ninguna tabla encontrada.", fg = "red")
        
        except Exception as e:
            messagebox.showerror("Error," f"No se pudo leer el PDF: \n{e}")
            #self.lbl_estado_pdf.config(text="Error", fg="red")

    def escanear_archivos(self):
        try:
            enlaces = self.driver.find_elements(By.TAG_NAME, "a")
            self.enlaces_encontrados= [] #Limpia lo que se encontró antes.
            
            ext_interes = ('.csv', '.xlsx', '.zip', '.pdf')             #Exensiones de interés MODIFICABLE 

            for widget in self.scroll_frame.winfo_children():           #Limpiar interfaz
                widget.destroy()
            self.check_vars = []

            contador = 0                                            
            for link in enlaces:                                        #Búsqueda de info
                try:
                    url = link.get_attribute('href')
                    texto = link.text.strip()                           #Condicional para ver si es algo descargable 
                    
                    if url and (url.lower().endswith(ext_interes) or "download" in url.lower() or "descargar" in texto.lower()):
                        if not texto:                                   #SI no hay link, se busca en el titulo 
                                texto = os.path.basename(url)
                        self.enlaces_encontrados.append((texto, url))  #Se guarda la info
                        var = tk.BooleanVar()
                        chk = tk.Checkbutton(self.scroll_frame, text = f"{texto} \n({url.split('/')[-1][:30]}...)", variable=var, bg="white", justify="left", anchor="w")
                        chk.pack(fill="x", pady=2)
                        self.check_vars.append(var)
                        contador += 1
                except:         
                    continue                       #Si falla, no se detiene la ejecución
        
            self.driver.quit()
            self.driver = None

            if contador > 0:
                self.lbl_estado_files.config(text = f"Se encontraron {contador} archivos. \n Selecciona y descarga.", fg = "green")
                self.btn_download_files.config(state = "normal")
            else:
                self.lbl_estado_files.config(text="No se encontraron archivos para descargar.", fg= "red")

        except Exception as e:
            messagebox.showerror("Error de búsqueda.", f"{e}")


    def descargar_archivos_seleccionados(self):                                                # Función para descargar los archivos que eligió el usuario
        indices_seleccionados = [i for i, var in enumerate(self.check_vars) if var.get()]    #Obtener los archivos seleccionados
        if not indices_seleccionados:
            messagebox.showwarning("Cuidado", "No seleccionaste ningún archivo")
            return
        carpeta_destino = filedialog.askdirectory(title="Selecciona la carpeta en donde guardar los archivos")   #Carpeta destino
        if not carpeta_destino: return
        exito  = 0
        errores = 0
        self.lbl_estado_files.config(text="Descargando. . .", fg ="blue")
        self.root.update()

        for i in indices_seleccionados:                             #Descargar archivos uno por uno
            nombre_display, url = self.enlaces_encontrados[i]

            try: 
                nombre_archivo = url.split('/')[-1]   #Se intenta buscar un nombre para el archivo
                if "." not in nombre_archivo:
                    nombre_archivo += ".csv"     #Si no tiene extensión en el NOMBRE, se le añade
                ruta_completa  =os.path.join(carpeta_destino, nombre_archivo)

                response  =requests.get(url, stream=True)
                if response.status_code == 200:
                    with open(ruta_completa, 'wb') as f:
                        for chunk in response.iter_content(chunk_size=1024):
                            f.write(chunk)
                    exito +=1
                else:
                    errores += 1
            except:
                errores += 1

        messagebox.showinfo("Resultados", f"Descargas realizadas: {exito} de forma exitosa.\n Errores: {errores}.\n Se guardaron en: {carpeta_destino}.")
        self.lbl_estado_files.config(text="Proceso finalizado", fg="green")

    def guardar_excel_csv(self, tablas, origen):
        archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
        if not archivo: return

        try:
            if archivo.endswith(".xlsx"):               # t = tablas, c = columnas
                with pd.ExcelWriter(archivo) as writer: #exclewriter es para que cada tabla descargada sea una hoja de excel 
                    for i, t in enumerate(tablas):      #Ciclo for para iterar en cada tabla
                        # Aplanar MultiIndex            #A cada tabla se le da el nombre 'Tabla i'
                        if isinstance(t.columns, pd.MultiIndex):  #se añade esto para 'aplanar' las celdas combinada porque si no, da error.
                            t.columns = [' '.join(map(str, c)).strip() for c in t.columns.values]
                        t.to_excel(writer, sheet_name=f"Tabla_{i+1}", index=False) # Guardamos cada tabla en su hoja correspondiente
            
            elif archivo.endswith(".csv"):
                base = os.path.splitext(archivo)[0]
                for i, t in enumerate(tablas):
                    if isinstance(t.columns, pd.MultiIndex):
                            t.columns = [' '.join(map(str, c)).strip() for c in t.columns.values]
                    
                    nombre = f"{base}_Tabla_{i+1}.csv" if len(tablas)>1 else archivo
                    t.to_csv(nombre, index=False, encoding='utf-8-sig')

            messagebox.showinfo("Éxito", f"Guardado correctamente desde {origen}.")
        except Exception as e:
            messagebox.showerror("Error Guardado", f"{e}")

#Bloque de ejecución
if __name__ == "__main__":
    ventana = tk.Tk()
    app = ScraperApp(ventana)
    ventana.mainloop()