""" 
Se trabaja en la pestaña para convertir archivos xls o xlsx a csv y viceversa
 F U N C I O N A L 
"""
import pandas as pd                                               #Para trabajar con datos
import requests
import os
import re                                                         #Para buscar patrones de números para crear los directorios
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog
from PIL import Image, ImageTk                                    #Para poner la imagen
import pdfplumber                                                 #Para trabajar con PDF
from io import StringIO                                           #Para tratar cadenas de texto como archivos
from selenium import webdriver                                    #Selenium es para poder trabajar con JavaScripts en la web
from selenium.webdriver.common.by import By                       #Para hacer búsquedas
from selenium.webdriver.chrome.service import Service       
from webdriver_manager.chrome import ChromeDriverManager          #Estas 3 líneas son para hacer búsquedas sin que nos bloquee la red por automatizarlo
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait           #Se cambia TIME por esta para esperar lo necesario para salvar esta compañía
from selenium.webdriver.support import expected_conditions as EC  #Condiciones de espera
from urllib.parse import unquote                                  #Librería para limpiar el nombre
import threading                                                  #Se meten hilos para que no se congele la ventana al buscar
import webbrowser                                                 #Para abrir enlaces OneDrive en el navegador
from datetime import datetime
import jinja2

class ScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Descargador de Datos")                   #Ventana principal
        self.root.geometry("700x800")                             #Dimensiones 
        self.driver = None

        self.check_vars  = []
        self.enlaces_encontrados = []
        self.enlaces_visibles = []
        self.urls_seleccionadas = set()                           #Para guardar los arch selecc. y no se borren al quitar los filtros
        self.ultima_url_buscada = ""                              #Para no hacer doble búsqueda con el mismo link 
        
        self.mis_cookies = [] 
        self.user_agent_usado = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

        self.icon_sesnsp = None                                   #Icono de la imagen sesnsp
        self.icon_datos = None                                    #Icono de la imagen datos 
        try:
            image_path = "reportes_incidencia.jpeg"               #Busca imagen para el botón 1
            if os.path.exists(image_path):
                img_original = Image.open(image_path)
                img_resized = img_original.resize((100,50), Image.Resampling.LANCZOS)
                self.icon_sesnsp = ImageTk.PhotoImage(img_resized)
        except Exception as e:
            print(f"No se pudo cargar la imagen")
        
        try:
            image_path_datos = "datos_abiertos.jpeg"              #Busca imagen para el botón 2
            if os.path.exists(image_path_datos):
                img_og_datos = Image.open(image_path_datos)
                img_res_datos = img_og_datos.resize((100,50), Image.Resampling.LANCZOS)
                self.icon_datos = ImageTk.PhotoImage(img_res_datos)
        except Exception as e:
            print(f"No se pudo cargar la imagen")

        #Separación de funciones por pestañas
        self.tab_control = ttk.Notebook(root)               

        self.tab_files = tk.Frame(self.tab_control)      
        self.tab_control.add(self.tab_files, text= ' Archivos encontrados ')            #Pestaña 1: Descargar archivos

        self.tab_verificador = tk.Frame(self.tab_control)      
        self.tab_control.add(self.tab_verificador, text= ' Verificador de enlaces ')    #Pestaña 2: Descargar archivos

        self.tab_conversor = tk.Frame(self.tab_control)      
        self.tab_control.add(self.tab_conversor, text= ' Conversor de archivos ')       #Pestaña 3: Descargar archivos

        self.tab_web = tk.Frame(self.tab_control)        
        self.tab_control.add(self.tab_web, text= ' Tablas en sitios Web ')              #Pestaña 4: Descargar tablas en sitios web

        self.tab_pdf = tk.Frame(self.tab_control)        
        self.tab_control.add(self.tab_pdf, text= ' Archivos PDF ')                      #Pestaña 5: Descargar tablas PDF de forma local

        self.tab_control.pack(expand=1, fill="both")


        #|||| 1 |||| PESTAÑA PARA BUSCAR Y DESCARGAR ARCHIVOS
        tk.Label(self.tab_files, text= " Buscar y descargar archivos (CSV, XLSX, PDF, ZIP Y RAR)", font = ("Arial", 14, "bold")).pack(pady=(15,5))
        tk.Label(self.tab_files, text= " Ingresa el URL de donde quieras descargar los datos", font = ("Arial", 10)).pack(pady=(8,5))

        #CAJA DE TEXTO (URL) + BOTÓN X
        frame_input = tk.Frame(self.tab_files)
        frame_input.pack(pady=5)

        self.entry_url_files = tk.Entry(frame_input, width=75)
        self.entry_url_files.pack(side="left", padx=(0, 5))                                        
        self.entry_url_files.bind('<Return>', lambda event: self.iniciar_hilo_busqueda("ARCHIVOS")) #Modo (archivos o tablas web)
        self.entry_url_files.focus()                                                                #Al abrir el programa, se puede escribir acá sin darle click necesariamente

        #BOTÓN X
        tk.Button(frame_input, text="X", command=self.limpiar_caja_url, bg="#FFCDD2", fg="red", font=("Arial", 8, "bold"), cursor="hand2", width=3).pack(side="left") 

        frame_botones_search = tk.Frame(self.tab_files)
        frame_botones_search.pack(pady=5)

        #BOTÓN BUSCAR ARCHIVOS
        tk.Button(frame_botones_search, text=" Buscar archivos ", command= lambda: self.iniciar_hilo_busqueda("ARCHIVOS"), bg = "#673AB7", fg="white").pack(side="left", pady=5)     
        
        #BOTÓN DE ACCESO DIRECTO A SESNSP
        if self.icon_sesnsp:
            btn_sesnsp = tk.Button(frame_botones_search, text="Incidencia delictiva", image= self.icon_sesnsp, compound="top", command=self.acceso_directo_sesnsp, cursor="hand2")
            btn_sesnsp.pack(side="left", padx=10)
        else:
            tk.Button(frame_botones_search, text="Ir a SESNSP", command= self.acceso_directo_sesnsp, bg="green", fg= "white").pack(side="left", padx=10)
        
        #BOTÓN DE DATOS ABIERTOS 
        if self.icon_datos:
            btn_datos = tk.Button(frame_botones_search, text="Datos Abiertos", image= self.icon_datos, compound="top", command= self.acceso_datos_abiertos, cursor="hand2")
            btn_datos.pack(side="left", padx=10)
        else:
            tk.Button(frame_botones_search, text="Datos Abiertos", command= self.acceso_datos_abiertos, bg ="green", fg="white").pack(side="left", padx=10)

        self.lbl_estado_files = tk.Label(self.tab_files, text= " Esperando URL . . .", fg="gray")
        self.lbl_estado_files.pack(pady=5)

        #FILTROS DE BÚSQUEDA:
        frame_filtros = tk.LabelFrame(self.tab_files, text= "Filtros ( Año / Palabra Clave)", padx= 5, pady = 5)
        frame_filtros.pack(fill="x", padx=20, pady=5)
        
        #FILTRO DE AÑO
        tk.Label(frame_filtros, text= " Año ").pack(side = "left", padx= 2)

        self.combo_filtro_anio = ttk.Combobox(frame_filtros, values=["Todos"], width=8, state= "readonly")
        self.combo_filtro_anio.current(0)
        self.combo_filtro_anio.pack(side = "left", padx=2)
        self.combo_filtro_anio.bind("<<ComboboxSelected>>", lambda event: self.aplicar_filtros())              #Aplica el filtro dando enter

        #FILTRO DE TEXTO
        tk.Label(frame_filtros, text= " Texto ").pack(side = "left", padx= 2)       
        self.entry_filtro_texto = tk.Entry(frame_filtros, width=15)
        self.entry_filtro_texto.pack(side = "left", padx=2)
        self.entry_filtro_texto.bind('<KeyRelease>', lambda event: self.aplicar_filtros())             #Aplica el filtro dando enter

        #FILTRO DE EXTENSIÓN
        tk.Label(frame_filtros, text= " Tipo de Archivo: ").pack(side ="left", padx=2)
        self.combo_ext = ttk.Combobox(frame_filtros, values =["Todos", ".csv", ".xlsx", ".pdf", ".zip", ".rar"], width= 7, state="readonly")
        self.combo_ext.current(0)                                                                  #"Todos" es marcado default
        self.combo_ext.pack(side="left", padx=2)
        self.combo_ext.bind("<<ComboboxSelected>>", lambda event: self.aplicar_filtros())          #Filtra con tan solo cliquear

        tk.Button(frame_filtros, text="Limpiar filtros", command=self.limpiar_filtros, bg = "#DDDDDD" ).pack(side="left")
        
        #AREA DE SELECCIÓN Y CONTADOR
        frame_seleccion = tk.Frame(self.tab_files)
        frame_seleccion.pack(fill="x", padx=20, pady=5)

        #BOTÓN PARA DESELECCIONAR
        tk.Button(frame_seleccion, text="Deseleccionar", command=self.deseleccionar_todo, bg = "#f0f0f0", font =("Arial", 10)).pack(side="right", padx=10)

        #Contador
        self.lbl_contador = tk.Label(frame_seleccion, text="Seleccionados: 0", font=("Arial", 9), fg="blue")
        self.lbl_contador.pack(side="right")

        #Checkbutton para Seleccionar todos los archivos 
        self.var_todos = tk.BooleanVar()
        self.chk_todos = tk.Checkbutton(self.tab_files, text= "Seleccionar / Deseleccionar todos", variable=self.var_todos, command=self.toggle_seleccion_todos)
        self.chk_todos.pack(pady = 2, anchor="w", padx= 20)

        #Checkbutton para organizar las carpetas por año
        self.var_organizar_anio = tk.BooleanVar()
        self.chk_organizar = tk.Checkbutton(self.tab_files, text="ORGANIZAR DESCARGAS POR CARPETAS", variable=self.var_organizar_anio, fg= "blue")
        self.chk_organizar.pack(pady=2, anchor="w", padx=20)

        #Por si hay muchos archivos, se hace una lista con checkboxes para elegir los que se quieran descargar.
        self.frame_lista = tk.Frame(self.tab_files, bd = 1, relief="sunken")             #Parte gráfica con Tk
        self.frame_lista.pack(side="top", fill="both", expand=True, padx= 20, pady = 5)

        #Barras de desplazamiento Ejes X e Y
        self.scrollbar_y = tk.Scrollbar(self.frame_lista, orient="vertical")
        self.scrollbar_x = tk.Scrollbar(self.frame_lista, orient="horizontal")

        self.canvas = tk.Canvas(self.frame_lista, bg="white", yscrollcommand=self.scrollbar_y.set, xscrollcommand=self.scrollbar_x.set) 

        self.scrollbar_y.config(command=self.canvas.yview)
        self.scrollbar_x.config(command=self.canvas.xview)

        self.scroll_frame = tk.Frame(self.canvas, bg = "white")

        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0,0), window=self.scroll_frame, anchor="nw")
        self.scrollbar_y.pack(side="right", fill = "y")
        self.scrollbar_x.pack(side="bottom", fill = "x")
        self.canvas.pack(side = "left", fill="both", expand = True)

        def _on_mousewheel(event):                                                      #Función para que sirva la rueda en la lista de archivos
            bbox = self.canvas.bbox("all")
            if not bbox: return
            height_content = bbox[3] - bbox[1]
            height_canvas = self.canvas.winfo_height()
            if height_content > height_canvas:                                          #Solo scrollea si la lista es más grande que el recuadro
                self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        self.frame_lista.bind('<Enter>', lambda e: self.canvas.bind_all("<MouseWheel>", _on_mousewheel))   #La rueda solo sirve dentro del recuadro de la lista
        self.frame_lista.bind('<Leave>', lambda e: self.canvas.unbind_all("<MouseWheel>"))

        self.btn_download_files = tk.Button(self.tab_files, text ="Descargar archivos.", command=self.descargar_archivos_seleccionados, state= "disabled", bg="#00baae", fg="white")
        self.btn_download_files.pack(side="bottom", pady = 10)

        #|||| 2 |||| PESTAÑA DE VERIFICACIÓN DE LINKS
        tk.Label(self.tab_verificador, text="Verificador de enlaces: ", font=("Arial", 12, "bold")).pack(pady=(15,5))
        tk.Label(self.tab_verificador, text="Carga tu archivo.\n Se creará un nuevo archivo indicando cuáles enlaces funcionan.", font=("Arial", 10)).pack(pady=5)
        
        tk.Button(self.tab_verificador, text="Seleccionar archivo", command = self.verificar_enlaces, bg="#FF9800", fg="white", height=2).pack(pady=10)

        self.lbl_estado_verif = tk.Label(self.tab_verificador, text="Esperando archivo. . .", fg="gray")
        self.lbl_estado_verif.pack(pady=5)
        
        #|||| 3 |||| PESTAÑA DE CONVERSIÓN DE ARCHIVOS XLS A CSV
        tk.Label(self.tab_conversor, text="Conversor de ARCHIVOS CSV A EXCEL Y VICEVERSA: ", font=("Arial", 13, "bold")).pack(pady=(15,20))
        tk.Label(self.tab_conversor, text="Elige tu archivo a convertir: ", font=("Arial", 12, "bold")).pack(pady=(15,5))
        tk.Button(self.tab_conversor, text="Seleccionar archivo", command = self.convertir_archivos, bg="#FF9800", fg="white", height=2).pack(pady=10)

        #|||| 4 |||| PESTAÑA DE TABLAS EN SITIOS WEB
        tk.Label(self.tab_web, text = "Ingresa el link para descargar sus tablas.", font =("Arial", 12, "bold")).pack(pady=(15,5))          #Instrucciones
        self.entry_url_web = tk.Entry(self.tab_web, width=50)                                                                               #Caja de texto de URL
        self.entry_url_web.pack(pady=5)
        self.entry_url_web.bind('<Return>', lambda event: self.iniciar_hilo_busqueda("WEB")) 

        tk.Button(self.tab_web, text = "Buscar y verificar tablas", command=lambda: self.iniciar_hilo_busqueda("WEB"), bg="#4CAF50", fg="white").pack(pady=(5))       #Botón que abre el link
        self.lbl_estado_web = tk.Label(self.tab_web, text="Esperando URL...", fg="gray", wraplength=450)                                                        #Estiqueta de espera
        self.lbl_estado_web.pack(pady=5)        
        self.btn_descargar_web = tk.Button(self.tab_web, text="Verificado. Descargar Excel", command=self.descargar_tablas, state="disabled", bg="#008CBA", fg="white", height=2)
        self.btn_descargar_web.pack(pady=10)                                                                                                                    #Botón de descarga

        #|||| 5 |||| PESTAÑA DE TABLAS EN PDF's
        tk.Label(self.tab_pdf, text = "Selecciona el archivo PDF desde tu computadora: ", font=("Arial", 12, "bold")).pack(pady=(15,5))
        tk.Button(self.tab_pdf, text="Seleccionar", command = self.procesar_pdf, bg="#FF9800", fg="white", height=2).pack(pady=10)
        self.lbl_estado_pdf = tk.Label(self.tab_pdf, text="No se eligió ningún archivo", fg="gray")
        self.lbl_estado_pdf.pack(pady = 20)

    def limpiar_caja_url(self):
        self.entry_url_files.delete(0, tk.END)
        self.ultima_url_buscada = ""

    def acceso_directo_sesnsp(self):                                                                                                #Función para entrar al SESNSP
        url_objetivo = "https://www.gob.mx/sesnsp/documentos/historico-de-incidencia-delictiva-del-fuero-comun"                     #Enlace
        self.entry_url_files.delete(0, tk.END)
        self.entry_url_files.insert(0, url_objetivo)                                       #Sustituye el link que está por el de SESNP
        self.iniciar_hilo_busqueda("ARCHIVOS")                         
    
    def acceso_datos_abiertos(self):
        url_objetivo = "https://www.gob.mx/sesnsp/acciones-y-programas/datos-abiertos-de-incidencia-delictiva"                      #Función para datos abiertos
        self.entry_url_files.delete(0, tk.END)                                            #Sustituye el link 
        self.entry_url_files.insert(0, url_objetivo)
        self.iniciar_hilo_busqueda("ARCHIVOS")                                            #Modo de descarga automático

    def iniciar_hilo_busqueda(self, modo):                                                 #Función para trabajar con hilos 
        hilo = threading.Thread(target=self.abrir_navegador, args=(modo,), daemon=True)    #Se ejecuta en segundo plano y no congela la ventana  
        hilo.start()

    def abrir_navegador(self, modo):        
        url = ""
        lbl = None
        if modo == "WEB":
            url = self.entry_url_web.get().strip()                      
            lbl = self.lbl_estado_web
        else:
            url = self.entry_url_files.get().strip()                    
            lbl= self.lbl_estado_files
        
        if not url:
            messagebox.showwarning("Error", "URL inválida.")            
            return

        if modo == "ARCHIVOS":
            if url == self.ultima_url_buscada:
                messagebox.showinfo("Información", "Ya se están mostrando los resultados de esta URL.\nNo es necesario recargar.")
                return
            else:
                self.ultima_url_buscada = url                           

        try:
            chrome_options = Options()                                                                  #Configuraciones del navegador
            chrome_options.add_argument("--start-maximized")            
            chrome_options.add_argument("--no-sandbox")                 
            chrome_options.add_argument("--disable-dev-shm-usage")      
            chrome_options.add_argument("--disable-gpu") 
            chrome_options.add_argument("--window-size=1920,1080") 

            chrome_options.add_argument(f"user-agent={self.user_agent_usado}")

            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])                            
            chrome_options.add_experimental_option('useAutomationExtension', False)
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")

            if modo == "ARCHIVOS":                                                                                      
                chrome_options.add_argument("--headless=new")
                lbl.config(text="Buscando en segundo plano... ", fg="black") #Mensaje actualizado
            else:
                lbl.config(text="Iniciando navegador... Por favor espera.", fg="black")                                           

            #self.root.update()  #Como es un hilo, ya no se forza la actualización
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)    
            self.driver.get(url)                                                                                        

            if modo == "WEB":                                                                                           
                lbl.config(text="Navega en la ventana hasta tener las tablas que desees.\nCuando las tengas, presiona el botón de abajo.", fg="black")
                self.btn_descargar_web.config(state="normal")

            elif modo== "ARCHIVOS":                                                             
                self.escanear_archivos()

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"No se pudo abrir el navegador:\n{e}"))   #Si hay error, mostrarlo en el hilo principal
            self.ultima_url_buscada = "" 
            if self.driver: 
                try: self.driver.quit()
                except: pass

    def procesar_pdf(self):                                                                                     #Función para procesar archivos PDF
        archivo = filedialog.askopenfilename (title = "Seleccionar Archivo PDF.", filetypes=[("PDF", "*.pdf")])
        if not archivo:
            return
        self.lbl_estado_pdf.config(text=f"Procesando: {os.path.basename(archivo)}...", fg="blue")
        self.root.update()
        try:
            tablas_totales = []     
            with pdfplumber.open(archivo) as pdf:                                       #Se abre el pdf
                for pagina in pdf.pages:
                    extracto = pagina.extract_tables()                                  #Extrae tablas de la página actual

                    for t in extracto:
                        df = pd.DataFrame(t[1:], columns=t[0]).dropna(how='all')        #Convertir a Dataframe
                        if not df.empty:
                            tablas_totales.append(df)                                   #Elimina filas totalmente vacías
            if tablas_totales:
                self.guardar_excel_csv(tablas_totales, "PDF")
                self.lbl_estado_pdf.config(text="Proceso realizado con éxito.", fg="green")
            else:
                messagebox.showwarning("PDF","No se encontraron tablas en el archivo")
                self.lbl_estado_pdf.config(text="Ninguna tabla encontrada.", fg = "red")

        except Exception as e:
            messagebox.showerror("Error," f"No se pudo leer el PDF: \n{e}")
            
    def escanear_archivos(self):                                                                
        try:
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, "a"))              #Espera hasta 10s, pero si sale antes avanza
                )
            except:
                pass                                                                #Si pasa el tiempo, continuamos por si acaso ya estaba cargado

            enlaces = self.driver.find_elements(By.TAG_NAME, "a")
            self.enlaces_encontrados= [] 
            
            ext_interes = ('.csv', '.xlsx', '.zip', '.pdf', '.rar')                                             
            temp_encontrados = []
            contador = 0                                            

            for link in enlaces:
                try:
                    url = link.get_attribute('href')

                    if not url: 
                        continue
                    texto = link.text.strip()

                    es_nombre_generico = any(x in texto.lower() for x in ["descargar", "documento", "archivo", "download"]) 
                    es_nube = any(x in url.lower() for x in ["onedrive", "sharepoint", "idrv.ms", "go.microsoft"])
                    es_nombre_generico = any(x in texto.lower() for x in ["descargar", "documento","archivo", "download"])                    
                    
                    if es_nombre_generico:                                                      #Búsqueda de un mejor nombre            
                        try:
                            elemento_padre = link.find_element(By.XPATH, "./..")                           #Busca en ele. padre                       
                            texto_padre = elemento_padre.text.strip()                                   
                            if len(texto_padre) > len(texto) + 3:               
                                texto =  texto_padre.replace(texto, "").strip()             
                            else:
                                elemento_abuelo = link.find_element(By.XPATH, "./../..")                   #Si no hay, en el awelo
                                texto_abuelo = elemento_abuelo.text.strip()
                                if len(texto_abuelo) > len(texto) + 3:                              
                                    texto = texto_abuelo.replace(texto, "").strip()
                        except:
                            pass

                    cumple_condicion = (url.lower().endswith(ext_interes) or "download" in url.lower() or "descargar" in texto.lower() or es_nube)            #Condición de aceptación
                    
                    if cumple_condicion:
                        if not texto:
                                texto = unquote(os.path.basename(url))                                      
                        texto = texto.replace("\n", " ").replace("\r", "")                                 #Limpia saltos de linea

                        temp_encontrados.append((texto, url))                                       
                        contador += 1
                except:        
                    continue                                                                                

            self.mis_cookies = self.driver.get_cookies()
            self.driver.quit()
            self.driver = None
            
            self.enlaces_encontrados = temp_encontrados                                             #Guardamos en la variable de clase
            self.root.after(0, self.finalizar_busqueda_gui, contador)                               #Actualiza pantalla

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error de búsqueda.", f"{e}"))

    def finalizar_busqueda_gui(self, contador):                                                     #Actualiza la interfaz gráfica del usuario (GUI)
        for widget in self.scroll_frame.winfo_children():                                           #Limpia interfaz
            widget.destroy()

        self.check_vars = []
        self.urls_seleccionadas.clear()
        #<--- ESTA ES LA LÍNEA QUE SUPUESTAMENTE "FALTA" PERO QUE SÍ EXISTE --->
        self.enlaces_visibles = self.enlaces_encontrados[:]

        conjunto_anios = set()                                         #Actualiza el filtro Anio con base en los años encontrados
        for texto, url in self.enlaces_encontrados:
            encontrados = re.findall(r'\b(20\d{2}|19\d{2})\b', texto)  #Se buscan fechas en los archivos
            if encontrados:
                for a in encontrados:
                    conjunto_anios.add(a)

        lista_anios = sorted(list(conjunto_anios), reverse=True)        #Ordenamos de mayor a menor
        lista_anios.insert(0, "Todos")

        self.combo_filtro_anio['values'] = lista_anios                  #Se actualiza el combobox
        self.combo_filtro_anio.current(0)

        self.actualizar_lista_visual(self.enlaces_visibles)

        if contador > 0:
            self.lbl_estado_files.config(text = f"Total de Archivos: {contador}", fg = "green")     #Muestra N archivos encontrados
            self.btn_download_files.config(state = "normal")
            self.actualizar_contador()  
            self.canvas.yview_moveto(0)                                                             #Resetea el canvas                                                        
        else:
            self.lbl_estado_files.config(text="No se encontraron archivos para descargar.", fg= "red")
            for widget in self.scroll_frame.winfo_children():
                widget.destroy()
            respuesta = messagebox.askyesno("Sin resultados", "No se hallaron resultados. \n ¿Quieres volver a intentarlo?\n A veces la página puede tardar en responder") #Pregunta de reintento
            if respuesta:
                self.ultima_url_buscada =""
                self.iniciar_hilo_busqueda("ARCHIVOS")                  #Llama al hilo de nuevo

    def actualizar_contador(self):
        count = len(self.urls_seleccionadas)                                                #Muestra los N archivos seleccionados
        self.lbl_contador.config(text=f"Seleccionados: {count}")                            #Se actualiza automáticamente

    def deseleccionar_todo(self):
        self.urls_seleccionadas.clear()                                                     #Limpia y refresca
        self.var_todos.set(False)
        self.aplicar_filtros()                                                              #Re-aplica filtros para que se muestren en su orden original

    def gestionar_click_individual(self, url, var_asociada):
        if var_asociada.get():
            self.urls_seleccionadas.add(url)
        else:
            self.urls_seleccionadas.discard(url)
        self.aplicar_filtros()                                                              #Se llama a esta funcion pq ya incluye el ordenamiento
    

    def actualizar_lista_visual(self, lista_datos):                                     #Actualiza los archivos enlistados según los filtros
        for widget in self.scroll_frame.winfo_children(): 
            widget.destroy()                                                                #Limpia interfaces
        
        self.check_vars = []                                                                #Reinicia checkboxes
        self.enlaces_visibles = lista_datos

        for texto, url in lista_datos:
            var = tk.BooleanVar()
            
            if url in self.urls_seleccionadas:                                              #Si está seleccionado, se mantiene
                var.set(True)

            try: 
                if "onedrive" in url or "sharepoint" in url:                                #Cambio visual de nombre si halla un OneDrive
                    nombre_archivo_completo = "Enlace a Carpeta / Archivo (OneDrive)"
                else:
                    nombre_archivo_completo = unquote(url.split('/')[-1])
            except: 
                nombre_archivo_completo = "Archivo"
            bg_color = "white"
            if "onedrive" in url or "sharepoint" in url:  
                bg_color = "#E3F2FD" #Los cloudfiles se distinguen por el color
            frame_fila = tk.Frame(self.scroll_frame, bg =bg_color)
            frame_fila.pack(fill="x", pady=1)

            chk = tk.Checkbutton(self.scroll_frame, text = f"{texto}\n({nombre_archivo_completo})", 
                                 variable=var, bg="white", justify="left", anchor="w",
                                 command=lambda u=url, v=var: self.gestionar_click_individual(u, v))   #Pasamos el URL específico a la función lambda
            
            chk.pack(fill="x", pady = 2)
            self.check_vars.append(var)

        self.canvas.configure(scrollregion=self.canvas.bbox("all"))                     #Actualiza lista
        self.actualizar_contador()                                                      
    
    def toggle_seleccion_todos(self): 
        valor = self.var_todos.get()                                                        #Se selecciona todo
        
        for _, url in self.enlaces_visibles:                                                #Se actualiza la memoria con todo lo visible 
            if valor:
                self.urls_seleccionadas.add(url)
            else:
                self.urls_seleccionadas.discard(url)
        self.aplicar_filtros()                                                              #Se ordenan 

    def aplicar_filtros(self):
        filtro_anio = self.combo_filtro_anio.get()                                 #Obtenemos año
        if filtro_anio == "Todos":                                             #Si está en todos, no filtra
            filtro_anio = ""

        filtro_texto = self.entry_filtro_texto.get().strip().lower()                        #Obtenemos cadena
        filtro_ext = self.combo_ext.get()                                                   #Obtenemos extensión

        if not self.enlaces_encontrados: return                                             #Si no halla, continua

        lista_filtrada = []
        for texto, url in self.enlaces_encontrados:
            cumple_anio, cumple_texto, cumple_ext = True, True, True
            
            if filtro_anio:                                                                                                         #Se busca el año escrito
                if filtro_anio not in texto and filtro_anio not in url: cumple_anio  = False
            if filtro_texto:                                                                                                        #Se busca la palabra
                if filtro_texto not in texto.lower() and filtro_texto not in url.lower(): cumple_texto = False
            if filtro_ext and filtro_ext != "Todos":
                if not url.lower().endswith(filtro_ext): cumple_ext = False                                                         #Verifica la extensión

            if cumple_anio and cumple_texto and cumple_ext: 
                lista_filtrada.append((texto, url))
        
        lista_filtrada.sort(key=lambda x: x[1] in self.urls_seleccionadas, reverse=True)                                  #Se ordena la lista filtrada 

        self.actualizar_lista_visual(lista_filtrada)                                                                      #Se actualiza con base a la l_filtrada
        
        if filtro_ext == "Todos" and not filtro_anio and not filtro_texto:
            self.lbl_estado_files.config(text=f"Total de Archivos: {len(self.enlaces_encontrados)}", fg="green")    #Si no hay filtro, se muestra todo
        else:
            self.lbl_estado_files.config(text=f"Mostrando {len(lista_filtrada)} de {len(self.enlaces_encontrados)} encontrados.", fg = "blue") #Si sí hay, se muestran 'X' filtrados del 'A' total

    def limpiar_filtros (self):
        self.combo_filtro_anio.current(0)                                                       #Cambio a combobox
        self.entry_filtro_texto.delete(0, tk.END)                                               #Elimina ambos filtros
        self.combo_ext.current(0)                                                               #Vuelve a 'todos'
        self.aplicar_filtros()                                                                  #Reordena
        
    def descargar_archivos_seleccionados(self):
        urls_a_descargar = [url for url in self.urls_seleccionadas]                             #Obtenemos la lista real de URLs a descargar desde la MEMORIA
        cantidad = len(urls_a_descargar)                                                        #Se contempla la cantidad a descargar
        
        if cantidad == 0:
            messagebox.showwarning("Cuidado", "No seleccionaste ningún archivo")        
            return
        
        #Verificación en NUBE: Se abrirá en el navegador
        nube_links = [u for u in urls_a_descargar if any(x in u.lower() for x in ["onedrive", "sharepoint", "idrv.ms"])]
        descarga_directa = [u for u in urls_a_descargar if u not in nube_links]

        if nube_links:                                                                          #Mensaje advirtiendo que se abrirá el navegador
            mensaje = f"Has seleccionado {len(nube_links)} enlaces. \n\n"\
            "Estos archivos no se pueden descargar directamente.\n"\
            "Se abrirá el navegador para que descargues la carpeta seleccionada\n\n¿Continuar"
            if messagebox.askyesno("Enlcaes externos", mensaje):
                for link in nube_links:
                    webbrowser.open(link)

        if not descarga_directa:
            self.deseleccionar_todo()
            return
        
        urls_a_descargar = descarga_directa                                                     #Procesos solo directos
        cantidad = len(urls_a_descargar)

        organizar_por_anio = self.var_organizar_anio.get()                                      #Verifica este modo de descarga
        modo_descarga = "MANUAL"                                                                #Modo Default
        
        if organizar_por_anio:          
            modo_descarga = "AUTOMATICO"                                                        #Si se descarga por bibliotecas, se hace automático
        else:
            if cantidad == 1:                                                                   #Si es solo un archivo manda directo a descargar 
                modo_descarga == "MANUAL"
            else:
                if cantidad > 7:                                                                #LIMTIE ARBITRARIO DE ARCHIVOS MANUALES, PUEDE CAMBIARSE
                    modo_descarga = "AUTOMATICO"                                                #Si se pasa el límite ahorra molestias y descarga autom.
                else:
                    respuesta =  messagebox.askyesno("Descarga de Archivos", f"Has seleccionado {cantidad} archivos. \n\n ¿Deseas nombrarlos uno por uno?\n Si eliges NO, se descargarán automáticamente con su nombre original.")  #Si son pocos, igual puede evitar nombrarlos
                    if not respuesta: modo_descarga = "AUTOMATICO"              

        carpeta_destino = ""
        if modo_descarga == "AUTOMATICO":
            if not organizar_por_anio:                                                                                                                                              #Solo pregunta si pasó el límite de 7
                respuesta =  messagebox.askyesno("Descarga de Archivos", f"¿Deseas descargar {cantidad} archivos?\nPuede ser un proceso tardado.")   #Verificación
                if not respuesta: return
            carpeta_destino = filedialog.askdirectory(title = "Selecciona la carpeta donde quieres descargar tus archivos.")                         #Elegir carpeta
            if not carpeta_destino: return

        exito, errores = 0, 0
        self.lbl_estado_files.config(text="Descargando...", fg ="blue")
        self.root.update()

        session = requests.Session()
        for cookie in self.mis_cookies:
            session.cookies.set(cookie['name'], cookie['value'])

        for url in urls_a_descargar:                                                                     #Iteramos sobre las URLs guardadas en memoria
            nombre_display = "Archivo desconocido"                                                       #Buscamos el texto original para el nombre (opcional, pero útil para mostrar)
            for txt, u in self.enlaces_encontrados:
                if u == url:
                    nombre_display = txt
                    break
            try:
                nombre_sugerido = unquote(url.split('/')[-1]).split('?')[0]
                if "." not in nombre_sugerido: nombre_sugerido += ".csv"    
                
                if "sharepoint.com" in url or "onedrive.aspx" in url:
                    if "download=1" not in url:
                        url += "&download=1"

                ruta_completa = ""
                if organizar_por_anio:                                                                                                  #Condicional Descargar por Año
                    anios_encontrados = re.findall(r'\b(19\d{2}|20\d{2})\b', nombre_display + " " + nombre_sugerido)                    #Formato de búsqueda de carpetas
                    carpeta_anio = anios_encontrados[0] if anios_encontrados else "Otros_Sin_Fecha"                                     #Si se encuentra el año, se asigna
                    
                    ruta_directorio_anio = os.path.join(carpeta_destino, carpeta_anio)
                    os.makedirs(ruta_directorio_anio, exist_ok=True)                                                                                                
                    ruta_completa = os.path.join(ruta_directorio_anio, nombre_sugerido)                                                 #Asignación de archivos por carpeta
                
                elif modo_descarga == "MANUAL":                                                                                         #Descarga en modo Manual
                    _, extension_original = os.path.splitext(nombre_sugerido)                         
                    ruta_completa = filedialog.asksaveasfilename(title=f"Guardar: {nombre_display[:50]}...", initialfile=nombre_sugerido, filetypes=[("Todos", "*.*")])
                    if not ruta_completa: continue                                                                                      
                    if not os.path.splitext(ruta_completa)[1]: ruta_completa += extension_original                                      #Asignación del nombre del archivo + extensión
                else:                                   
                    ruta_completa = os.path.join(carpeta_destino, nombre_sugerido) 

                response  = session.get(url, stream=True, verify=False)

                if response.status_code == 200:
                    with open(ruta_completa, 'wb') as f:
                        for chunk in response.iter_content(chunk_size=8192):
                            f.write(chunk)
                    if os.path.getsize(ruta_completa) < 2000:
                        print(f"ALERTA: El archivo {nombre_sugerido} parece un error de acceso.")
                    exito +=1
                else:
                    errores += 1
            except Exception as e:
                print(e)
                errores += 1

        if exito > 0 or errores > 0:
            messagebox.showinfo("Resultados", f"Descargas exitosas: {exito}.\nErrores: {errores}.")          #Si falló o descargó al menos un archivo
        else:
            if modo_descarga == "MANUAL":
                messagebox.showinfo("Resultados", "No se guardó ningún archivo.")                            #Si se cancelaron todas las descargas

        self.lbl_estado_files.config(text="Proceso finalizado", fg="green")
        self.var_organizar_anio.set(False)
        self.deseleccionar_todo()                                                                            #Al descargar algo. Se limpia la pantalla
        self.limpiar_filtros()

    def verificar_enlaces(self):
        archivo_path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])       #Como sé cual archivo es, solo admito excel
        if not archivo_path: return

        try:
            df= pd.read_excel(archivo_path)
            if "Link" not in df.columns:
                messagebox.showerror("Error", "Columna de enlaces no encontrada.")
                return
            
            self.lbl_estado_verif.config(text="Verificando enlaces...", fg="blue")

            hilo = threading.Thread(target=self.procesar_verificacion, args=(df,), daemon=True)
            hilo.start()
        except Exception as e:
            messagebox.showerror("Error de archivo", f"No se pudo leer el archivo: \n {e}")

    def convertir_archivos(self):
        file_path = filedialog.askopenfilename(title="Seleccionar archivo para convertir", filetypes=[("Archivos de Datos", "*.xlsx *.xls *.csv")])
        if not file_path: return
        _, ext = os.path.splitext(file_path)
        ext = ext.lower()
        try:
            if ext in ['.xlsx', '.xls']:                                              #--- EXCEL A CSV ---
                dict_dfs = pd.read_excel(file_path, sheet_name=None)                                                                            #sheet_name=None lee las hojas y devuelve {NombreHoja: DataFrame}
                save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], title="Guardar como CSV")       #Pedimos nombre base para guardar
                
                if not save_path: return

                base, extension = os.path.splitext(save_path)
                archivos_generados = 0                                               #Iteramos sobre cada hoja encontrada
                for nombre_hoja, df in dict_dfs.items():
                    
                    if len(dict_dfs) > 1:                                            #Si hay más de una hoja, añadimos el nombre de la hoja al archivo
                        nombre_limpio = nombre_hoja.replace(" ", "_")                #Limpiamos el nombre de la hoja para evitar caracteres raros en la ruta
                        ruta_final = f"{base}_{nombre_limpio}{extension}"
                    else: 
                        ruta_final = save_path                                       #Si solo hay una hoja, usamos el nombre exacto que eligió el usuario
                    
                    df.to_csv(ruta_final, index=False, encoding='utf-8-sig')
                    archivos_generados += 1

                msg_extra = f"\nSe generaron {archivos_generados} archivos CSV correspondientes a las hojas del Excel." if len(dict_dfs) > 1 else ""
                messagebox.showinfo("Éxito", f"Conversión completada.{msg_extra}")

            elif ext == '.csv':                                                         #--- CSV A EXCEL ---
                try:
                    df = pd.read_csv(file_path, encoding='utf-8')
                except:
                    df = pd.read_csv(file_path, encoding='latin-1')
                
                save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], title="Guardar como Excel")
                if save_path:
                    with pd.ExcelWriter(save_path) as writer:                          #Usamos ExcelWriter para asegurar un guardado limpio
                        df.to_excel(writer, index=False, sheet_name="Datos_CSV")
                    messagebox.showinfo("Éxito", "Archivo convertido a Excel correctamente.")
            
            else:
                messagebox.showwarning("Formato no soportado", "Por favor selecciona un archivo .xls, .xlsx o .csv")

        except Exception as e:
            messagebox.showerror("Error de conversión", f"No se pudo convertir el archivo:\n{e}")

    def procesar_verificacion(self, df):
        links_guardados = []                                                        #Lista para los links únicos
        resultados = []                                                             #Lista para sus estados
        headers = {'User-Agent': self.user_agent_usado}
        df['Link'] = df['Link'].astype(str).str.strip()                             #Convertimos a string y quitamos espacios en blanco
        lista_unica = df['Link'].unique()                                           #Obtenemos los valores únicos del DataFrame

        for link in lista_unica:
            if link.lower() == 'nan' or link == "":                                 #Si es 'nan' o cadena vacía, lo saltamos directamente, quitando filas amarillas y vacías
                continue
            estado = "Invalido"
            try:
                r = requests.head(str(link), headers=headers, timeout=5, allow_redirects=True) 
                if r.status_code == 200:
                    estado = "Valido"
                elif r.status_code == 405:
                    r = requests.get(str(link), headers=headers, timeout=5, stream=True)
                    if r.status_code == 200:
                        estado = "Valido"
                else:
                    r = requests.get(str(link), headers=headers, timeout=5, stream=True)
                    if r.status_code == 200: 
                        estado = "Valido"
            except:
                estado = "Invalido"
            
            links_guardados.append(link)                                            #Agregamos a las listas finales
            resultados.append(estado)
        
        df_filtrado = pd.DataFrame({                                                #Creamos el DataFrame final con los datos únicos
            "Link Original": links_guardados, 
            "Estado": resultados
        })

        self.root.after(0, self.finalizar_verificacion, df_filtrado)                #Enviamos el resultado (ya filtrado y sin duplicados) a finalizar
    
    def finalizar_verificacion(self, df_resultado):
        self.lbl_estado_verif.config(text="Verificación completada.", fg="green")

        fecha_hoy = datetime.now().strftime("%d_%m_%Y")
        nombre_sugerido = f"Enlaces_Vigentes_{fecha_hoy}"
        
        archivo_guardar = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile= nombre_sugerido, filetypes=[("Excel", "*.xlsx")], title="Guardar reporte")
        if not archivo_guardar:
            return

        try:
            def color_celdas(val):
                color = ''
                if val == 'Valido':
                    color = 'background-color: #C6EFCE; color: #006100' #Verde
                elif val == 'Invalido':
                    color = 'background-color: #FFC7CE; color: #9C0006' #Rojo
                return color
            try:
                styler = df_resultado.style.map(color_celdas, subset=['Estado'])    #Aplicar estilos
                styler.to_excel(archivo_guardar, index=False)
            except ImportError:
                df_resultado.to_excel(archivo_guardar, index= False)
                messagebox.showwarning("Advertencia", "Se guardó sin colores (instalar librería jinja2)")
            
            messagebox.showinfo("Éxito", f"Reporte guardado en:\n {archivo_guardar}")
        except Exception as e:
            messagebox.showerror("Error al guardar", f"{e}")

    def guardar_excel_csv(self, tablas, origen):                                                             #Función para tablas sencillas
        archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
        if not archivo: return
        try:
            if archivo.endswith(".xlsx"):               
                with pd.ExcelWriter(archivo) as writer: 
                    for i, t in enumerate(tablas):                                                                                      #t = tablas, c= columnas
                        if isinstance(t.columns, pd.MultiIndex):                                                                        #Excelwriter es para que cada tabla sea una hoja de excel
                            t.columns = [' '.join(map(str, c)).strip() for c in t.columns.values]                                       #Ireca en cada tabla
                        t.to_excel(writer, sheet_name=f"Tabla_{i+1}", index=False)                                                      #Cada tabla de guarda en su hoja correspondiente
            elif archivo.endswith(".csv"):                                                                                              #Aplana Multiindex, cada tabla se guarda como tabla i
                base = os.path.splitext(archivo)[0]
                for i, t in enumerate(tablas):
                    if isinstance(t.columns, pd.MultiIndex):
                            t.columns = [' '.join(map(str, c)).strip() for c in t.columns.values]
                    nombre = f"{base}_Tabla_{i+1}.csv" if len(tablas)>1 else archivo
                    t.to_csv(nombre, index=False, encoding='utf-8-sig')
            messagebox.showinfo("Éxito", f"Guardado correctamente desde {origen}.")
        except Exception as e:
            messagebox.showerror("Error Guardado", f"{e}")

    def reiniciar_interfaz(self):                                                       #Para que no se bloqueen los botones
        self.btn_descargar_web.config(state="disabled")
        self.lbl_estado_web.config(text="Listo.", fg="green")
        if self.driver:
            try: self.driver.quit()
            except: pass
            self.driver = None

    def descargar_tablas(self):
        if not self.driver: return
        try:
            html = self.driver.page_source
            self.driver.quit()
            self.driver = None
            tablas = pd.read_html(StringIO(html))
            if not tablas:
                messagebox.showwarning("Aviso", "No se encontraron tablas.")
                return
            self.guardar_excel_csv(tablas, "WEB")
            self.lbl_estado_web.config(text="Listo", fg="green")
            self.btn_descargar_web.config(state="disabled")
        except Exception as e:
            messagebox.showerror("Error", f"{e}")
        finally:
            self.reiniciar_interfaz()

if __name__ == "__main__":                                                              #Bloque de ejecución
    ventana = tk.Tk()
    app = ScraperApp(ventana)
    ventana.mainloop()