""" Lo de guardar por carpetas funciona, igualmente se filtra por extensión, se activan los filtros y búsquedas con ENTER, Correción de alertas    // N O   F U N C I O N A L 
    Se añadió un contador de archivos seleccionados y un botón de deseleccionar todo.
"""
import pandas as pd                                         #Para trabajar con datos
import requests
import os
import re                                                   #Para buscar patrones de números para crear los directorios
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog
import pdfplumber                                           # Para trabajar con PDF
from io import StringIO                                     # Para tratar cadenas de texto como archivos
from selenium import webdriver                              # Selenium es para poder trabajar con JavaScripts en la web
from selenium.webdriver.common.by import By                 # Para hacer búsquedas
from selenium.webdriver.chrome.service import Service       
from webdriver_manager.chrome import ChromeDriverManager    #Estas 3 líneas son para hacer búsquedas sin que nos bloquee la red por automatizarlo
from selenium.webdriver.chrome.options import Options
from urllib.parse import unquote                            #Librería para limpiar el nombre (quitar el "%20" en algunos docs que se desargan)

class ScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Descargador de Datos")
        self.root.geometry("700x750")
        self.driver = None

        self.check_vars  = []
        self.enlaces_encontrados = []
        self.enlaces_visibles = []
        self.urls_seleccionadas = set()                                                   #Para guardar los archivos seleccionados y no se borren al quitar filtros
        self.mis_cookies = []
        self.user_agent_usado = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

        #Separación de funciones por pestañas
        self.tab_control = ttk.Notebook(root)                                             #Para que se distribuya mejor la infterfaz, se va a dividir en pestañas

        self.tab_files = tk.Frame(self.tab_control)                                       #PESTAÑA 1: Para descargar archivos
        self.tab_control.add(self.tab_files, text= ' Archivos encontrados ')

        self.tab_web = tk.Frame(self.tab_control)                                         #PESTAÑA 2: Para sitios web
        self.tab_control.add(self.tab_web, text= ' Tablas en sitios Web ')

        self.tab_pdf = tk.Frame(self.tab_control)                                         #PESTAÑA 3: Para archivos PDF de forma local
        self.tab_control.add(self.tab_pdf, text= ' Archivos PDF ')

        self.tab_control.pack(expand=1, fill="both")

        #\\\ 1 /// PESTAÑA PARA BUSCAR Y DESCARGAR ARCHIVOS
        tk.Label(self.tab_files, text= " Buscar y descargar archivos (CSV, XLSX, PDF, ZIP Y RAR)", font = ("Arial", 14, "bold")).pack(pady=(15,5))
        tk.Label(self.tab_files, text= " Ingresa el URL de donde quieras descargar los datos", font = ("Arial", 10)).pack(pady=(8,5))
        self.entry_url_files = tk.Entry(self.tab_files, width=75)
        self.entry_url_files.pack(pady=5)
        self.entry_url_files.bind('<Return>', lambda event: self.abrir_navegador("ARCHIVOS"))            #Para que sirva ENTER en el cuadro de URL
        self.entry_url_files.focus()                                                                #Al abrir el programa, se puede escribir acá sin darle click necesariamente


        #BOTÓN BUSCAR ARCHIVOS
        tk.Button(self.tab_files, text=" Buscar archivos ", command= lambda: self.abrir_navegador("ARCHIVOS"), bg = "#673AB7", fg="white").pack(pady=5)     #MODO: "ARCHIVOS"
        self.lbl_estado_files = tk.Label(self.tab_files, text= " Esperando URL . . .", fg="gray")
        self.lbl_estado_files.pack(pady=5)

        #FILTROS DE BÚSQUEDA:
        frame_filtros = tk.LabelFrame(self.tab_files, text= "Filtros ( Año / Palabra Clave)", padx= 5, pady = 5)
        frame_filtros.pack(fill="x", padx=20, pady=5)
        
        #FILTRO DE AÑO
        tk.Label(frame_filtros, text= " Año ").pack(side = "left", padx= 2)
        self.entry_filtro_anio = tk.Entry(frame_filtros, width=8)
        self.entry_filtro_anio.pack(side = "left", padx=2)
        self.entry_filtro_anio.bind('<Return>', lambda event: self.aplicar_filtros())           #Para que sirva el ENTER 

        #FILTRO DE TEXTO
        tk.Label(frame_filtros, text= " Texto ").pack(side = "left", padx= 2)       
        self.entry_filtro_texto = tk.Entry(frame_filtros, width=15)
        self.entry_filtro_texto.pack(side = "left", padx=2)
        self.entry_filtro_texto.bind('<Return>', lambda event: self.aplicar_filtros())          #Para que sirva el ENTER

        #FILTRO DE EXTENSIÓN
        tk.Label(frame_filtros, text= " Tipo de Archivo: ").pack(side ="left", padx=2)
        self.combo_ext = ttk.Combobox(frame_filtros, values =["Todos", ".csv", ".xlsx", ".pdf", ".zip", ".rar"], width= 7, state="readonly")
        self.combo_ext.current(0)                                                               #"Todos" marcado por defecto
        self.combo_ext.pack(side="left", padx=2)
        self.combo_ext.bind("<<ComboboxSelected>>", lambda event: self.aplicar_filtros())       #Filtra con tan solo seleccionar

        tk.Button(frame_filtros, text="Aplicar filtros", command= self.aplicar_filtros, bg="#DDDDDD").pack(side="left", padx=10)
        tk.Button(frame_filtros, text="Quitar filtros", command=self.limpiar_filtros, bg = "#DDDDDD" ).pack(side="left")
        
        #AREA DE SELECCIÓN Y CONTADOR
        frame_seleccion = tk.Frame(self.tab_files)
        frame_seleccion.pack(fill="x", padx=20, pady=5)

        #Checkbutton para Seleccionar todos los archivos 
        self.var_todos = tk.BooleanVar()
        self.chk_todos = tk.Checkbutton(self.tab_files, text= "Seleccionar / Deseleccionar todos", variable=self.var_todos, command=self.toggle_seleccion_todos)
        self.chk_todos.pack(pady = 2, anchor="w", padx= 20)

        #BOTÓN PARA DESELECCIONAR
        tk.Button(frame_seleccion, text="Deseleccionar", command=self.deseleccionar_todo, bg = "#f0f0f0", font =("Arial", 10)).pack(side="right", padx=10)

        #Contador
        self.lbl_contador = tk.Label(frame_seleccion, text="Seleccionados: 0", font=("Arial", 9), fg="blue")
        self.lbl_contador.pack(side="right")

        #Checkbutton para organizar las carpetas por año
        self.var_organizar_anio = tk.BooleanVar()
        self.chk_organizar = tk.Checkbutton(self.tab_files, text="ORGANIZAR DESCARGAS POR CARPETAS", variable=self.var_organizar_anio, fg= "blue")
        self.chk_organizar.pack(pady=2, anchor="w", padx=20)

        #Por si hay muchos archivos, se hace una lista con checkboxes para elegir los que se quieran descargar.
        self.frame_lista = tk.Frame(self.tab_files, bd = 1, relief="sunken")               #Parte gráfica con Tk
        self.frame_lista.pack(side="top", fill="x", expand=True, padx= 20, pady = 5)

        self.canvas = tk.Canvas(self.frame_lista, bg="white")
        self.scrollbar = tk.Scrollbar(self.frame_lista, orient= "vertical", command=self.canvas.yview)
        self.scroll_frame = tk.Frame(self.canvas, bg = "white")

        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0,0), window=self.scroll_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side = "left", fill="both", expand = True)
        self.scrollbar.pack(side="right", fill = "y")

        #Código para que funcione la rueda del mouse (fuentes: chatgpt)
        def _on_mousewheel(event):                                                          # Obtenemos las dimensiones para saber si es necesario scrollear
            bbox = self.canvas.bbox("all")
            if not bbox: return
            height_content = bbox[3] - bbox[1]
            height_canvas = self.canvas.winfo_height()
            if height_content > height_canvas:                                              # Solo permitimos scroll si el contenido es más grande que la ventana
                self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        # Vinculamos la rueda solo cuando el mouse entra a la lista (para no afectar otras áreas)
        self.frame_lista.bind('<Enter>', lambda e: self.canvas.bind_all("<MouseWheel>", _on_mousewheel))
        self.frame_lista.bind('<Leave>', lambda e: self.canvas.unbind_all("<MouseWheel>"))

        self.btn_download_files = tk.Button(self.tab_files, text ="Descargar archivos.", command=self.descargar_archivos_seleccionados, state= "disabled", bg="#00baae", fg="white")
        self.btn_download_files.pack(side="bottom", pady = 10)

        #\\\ 2 /// PESTAÑA DE TABLAS EN SITIOS WEB
        tk.Label(self.tab_web, text = "Ingresa el link para descargar sus tablas.", font =("Arial", 12, "bold")).pack(pady=(15,5))          #instrucciones
        self.entry_url_web = tk.Entry(self.tab_web, width=50)                                                                               #URL caja de texto
        self.entry_url_web.pack(pady=5)
        self.entry_url_web.bind('<Return>', lambda event: self.abrir_navegador("WEB"))            #Para que sirva ENTER en el cuadro de URL

        #MODO: "WEB"                                                                                          
        tk.Button(self.tab_web, text = "Buscar y verificar tablas", command=lambda: self.abrir_navegador("WEB"), bg="#4CAF50", fg="white").pack(pady=(5))   #Boton para abrir el link
        self.lbl_estado_web = tk.Label(self.tab_web, text="Esperando URL...", fg="gray", wraplength=450)                                                           #Etiqueta de estado esperando
        self.lbl_estado_web.pack(pady=5)
        self.btn_descargar_web = tk.Button(self.tab_web, text="Verificado. Descargar Excel", command=self.descargar_tablas, state="disabled", bg="#00baae", fg="white", height=2)
        self.btn_descargar_web.pack(pady=10)       #Botón para descargar

        #\\\ 3 /// PESTAÑA DE TABLAS EN PDF's
        tk.Label(self.tab_pdf, text = "Selecciona el archivo PDF desde tu computadora: ", font=("Arial", 12, "bold")).pack(pady=(15,5))
        tk.Button(self.tab_pdf, text="Seleccionar", command = self.procesar_pdf, bg="#FF9800", fg="white", height=2).pack(pady=10)
        self.lbl_estado_pdf = tk.Label(self.tab_pdf, text="No se eligió ningún archivo", fg="gray")
        self.lbl_estado_pdf.pack(pady = 20)


    #En esta función se implementa el modo, por si es busqueda WEB o descarga de ARCHIVOS
    def abrir_navegador(self, modo):                                                                                          #FUNCION Para abrir el chrome                              
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
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            
            chrome_options.add_argument("--disable-gpu") # Necesario para headless en windows
            chrome_options.add_argument("--window-size=1920,1080") # Evita errores de renderizado 

            user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"  #Esto evita que se detecte como script automático
            chrome_options.add_argument(f"user-agent={user_agent}")                                                                         #Así, se previene que se nos bloqueé el acceso y muestre 0 archivos encontrados

            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])  #Esto es algo anti bots para intentar entrar a links más protegidos
            chrome_options.add_experimental_option('useAutomationExtension', False)
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")

            if modo == "ARCHIVOS":                                                                        #Apartado para abrir solo el navegador en modo web
                chrome_options.add_argument("--headless=new")
                lbl.config(text="Buscando en segundo plano... Por favor espera.", fg="black")             # No iniciamos el driver, se busca en 2do plano
            else:                                                                                         #Cuando es WEB
                lbl.config(text="Iniciando... Por favor espera.", fg="black")                             # Iniciamos el driver

            self.root.update()                                                                                          # Forzar actualización de la interfaz visual
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)     #Se incia el navegador
            self.driver.get(url)                                                                                                                        #Se obtiene el link

            if modo == "WEB":
                lbl.config(text="Navega en la ventana hasta tener las tablas que desees.\nCuando las tengas, presiona el botón de abajo.", fg="black")
                self.btn_descargar_web.config(state="normal")

            elif modo== "ARCHIVOS":
                lbl.config(text="Buscando archivos en la página. . .", fg="black")
                self.root.update()
                self.escanear_archivos()                                                                                #Función para escanear archivos

        except Exception as e:                                                                                            #Excepción por si hay error
            messagebox.showerror("Error", f"No se pudo abrir el navegador:\n{e}")
            if self.driver: self.driver.quit()

    def descargar_tablas(self):                                                                                                #Función para descargar las tablas
        if not self.driver:
            return
        try:
            html = self.driver.page_source            #Obtenemos el html completo
            self.driver.quit()                        #Se cierra el navegador
            self.driver = None                        
        
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

    def procesar_pdf(self):                                              #Función para procesar archivos PDF
        archivo = filedialog.askopenfilename (title = "Seleccionar Archivo PDF.", filetypes=[("PDF", "*.pdf")])
        if not archivo:
            return
        self.lbl_estado_pdf.config(text=f"Procesando: {os.path.basename(archivo)}...", fg="blue")
        self.root.update()
        try:
            tablas_totales = []     
            with pdfplumber.open(archivo) as pdf:                         #Se abre el pdf
                for pagina in pdf.pages:
                    extracto = pagina.extract_tables()                    #Extrae tablas de la página actual

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
            enlaces = self.driver.find_elements(By.TAG_NAME, "a")
            self.enlaces_encontrados= []                                #Limpia lo que se encontró antes.
            
            ext_interes = ('.csv', '.xlsx', '.pdf', '.zip', '.rar')                     #Exensiones de interés MODIFICABLE /// AÑADÍ RAR Y ZIP PORQUE NUNCA SE SABE

            for widget in self.scroll_frame.winfo_children():           #Limpiar interfaz
                widget.destroy()

            self.check_vars = []
            self.urls_seleccionadas.clear()                             #Se limpia la memoria en cada NUEVA búsqueda
            contador = 0                                    

            for link in enlaces:                                                             #Búsqueda de info
                try:
                    url = link.get_attribute('href')
                    texto = link.text.strip()                                                #Condicional para ver si es algo descargable

                    #Fragmento para que muestre el nombre completo del archivo y no un "Descargar documento"
                    es_nombre_generico = any(x in texto.lower() for x in ["descargar", "documento", "archivo", "download"])

                    if es_nombre_generico:
                        try:
                            elemento_padre = link.find_element(By.XPATH, "./..")             #Se busca el elemento padre 
                            texto_padre = elemento_padre.text.strip()

                            if len(texto_padre) > len(texto) + 3:                            # Si el padre tiene más texto que solo "Descargar"
                                texto =  texto_padre.replace(texto, "").strip()              #Se quita la palabra Descargar para que salga clean 
                            else:
                                elemento_abuelo = link.find_element(By.XPATH, "./../..")     #Si no se puso obtener el texto padre, se sube de nivel
                                texto_abuelo = elemento_abuelo.text.strip()
                                if len(texto_abuelo) > len(texto) + 3:
                                    texto = texto_abuelo.replace(texto, "").strip()
                        except:
                            pass

                    if url and (url.lower().endswith(ext_interes) or "download" in url.lower() or "descargar" in texto.lower()):
                        if not texto:                                                        #SI no hay link, se busca en el titulo
                                texto = unquote(os.path.basename(url))                       #Se limpia el nombre 

                        self.enlaces_encontrados.append((texto, url))                        #Se guarda la info
                        var = tk.BooleanVar()

                        nombre_archivo_completo = unquote(url.split('/')[-1])                #Con esto se marca el nombre del archivo completo para saber cuál descarg

                        chk = tk.Checkbutton(self.scroll_frame, text = f"{texto} \n({nombre_archivo_completo})", variable=var, bg="white", justify="left", anchor="w", command=self.actualizar_contador)
                        chk.pack(fill="x", pady=2)

                        self.check_vars.append(var)
                        contador += 1
                except:        
                    continue   #Si falla, no se detiene la ejecución
            
            self.mis_cookies = self.driver.get_cookies()

            self.driver.quit()
            self.driver = None
            # <--- ESTA ES LA LÍNEA QUE SUPUESTAMENTE "FALTA" PERO QUE SÍ EXISTE --->
            self.enlaces_visibles = self.enlaces_encontrados[:]

            if contador > 0:
                self.lbl_estado_files.config(text = f"Se encontraron {contador} archivos. \n Selecciona y descarga.", fg = "green")
                self.btn_download_files.config(state = "normal")
                self.actualizar_contador()
                self.canvas.yview_moveto(0)                                     #Mueve el canvas al inicio

            else:
                self.lbl_estado_files.config(text="No se encontraron archivos para descargar.", fg= "red")
                for widget in self.scroll_frame.winfo_children():
                    widget.destroy()

        except Exception as e:
            messagebox.showerror("Error de búsqueda.", f"{e}")

    def actualizar_contador(self):                                               #Se muestran los N archivos seleccionados para llevar una cuenta
        count= sum([var.get() for var in self.check_vars])
        self.lbl_contador.config(text=f"Seleccionados: {count}")                 #Se actualiza automáticamente 

    def deseleccionar_todo(self):
        for var in self.check_vars:
            var.set(False)
            self.var_todos.set(False)
            self.actualizar_contador()

    def actualizar_lista_visual(self, lista_datos):                               #Actualiza los archivos enlistados en la búsqueda de Files dependiendo el filtro      
        for i, var in enumerate(self.check_vars):                                 #Antes de borrar, se guardan los archivos seleccionados si es que hay
            if i < len(self.enlaces_visibles):
                url_actual = self.enlaces_visibles[i][1]
                if var.get():
                    self.urls_seleccionadas.add(url_actual)
                else:
                    self.urls_seleccionadas.discard(url_actual)

        for widget in self.scroll_frame.winfo_children():                         #Limpia interfaz
            widget.destroy()

        self.check_vars = []                                                      #Reinicia checkboxes
        self.enlaces_visibles = lista_datos

        for texto, url in lista_datos:
            var  = tk.BooleanVar()
            if url in self.urls_seleccionadas:
                var.set(True)
            try:
                nombre_archivo_completo = unquote(url.split('/')[-1])
            except:
                nombre_archivo_completo = "Archivo"

            chk = tk.Checkbutton(self.scroll_frame, text = f"{texto} \n({nombre_archivo_completo})", variable=var, bg="white", justify="left", anchor="w", command= self.actualizar_contador)
            chk.pack(fill="x", pady = 2)
            self.check_vars.append(var)

        self.var_todos.set(False)                                                           #Se resetean todos los checks para que no se queden marcados
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))                         #Actualizar lista
        self.actualizar_contador()                                                          #Afecta al contador
        self.canvas.yview_moveto(0)                                                         #Hace que si se buscan dos links seguidos, no se bugueé y salga en blanco (aunque los datos están más arriba)


    def toggle_seleccion_todos (self):                                                      #Función para select todos
        valor =self.var_todos.get()
        for var in self.check_vars:
            var.set(valor)
        self.actualizar_contador()                                                          #Afecta al contador

    def aplicar_filtros(self):
        filtro_anio = self.entry_filtro_anio.get().strip()                                  #Obtener año
        filtro_texto = self.entry_filtro_texto.get().strip().lower()                        #Obtener cadena
        filtro_ext = self.combo_ext.get()                                                   #Obtener extensión

        if not self.enlaces_encontrados:                                                    #Si no se hallan links ps no hace nada
            return

        lista_filtrada = []
        for texto, url in self.enlaces_encontrados:
            cumple_anio = True
            cumple_texto = True
            cumple_ext = True
            
            if filtro_anio:
                if filtro_anio not in texto and filtro_anio not in url:                     #Se busca que esté el año escrito
                    cumple_anio  = False
            
            if filtro_texto:
                if filtro_texto not in texto.lower() and filtro_texto not in url.lower():   #Se busca que esté la palabra
                    cumple_texto = False

            if filtro_ext and filtro_ext != "Todos":
                if not url.lower().endswith(filtro_ext):                                    #Verificar si acaba con la extension
                    cumple_ext = False

            if cumple_anio and cumple_texto and cumple_ext:
                lista_filtrada.append((texto, url))
        
        self.actualizar_lista_visual(lista_filtrada)
        #filtros_vacios = (not filtro_anio) and (not filtro_texto) and (filtro_ext == "Todos")
        if filtro_ext == "Todos" and (not filtro_anio) and (not filtro_texto):
            self.lbl_estado_files.config(text=f"Total de Archivos: {len(self.enlaces_encontrados)}", fg="green")
        else:
            self.lbl_estado_files.config(text=f"Mostrando {len(lista_filtrada)} de {len(self.enlaces_encontrados)} encontrados. ", fg = "blue")

    def limpiar_filtros (self):
        self.entry_filtro_anio.delete(0, tk.END)
        self.entry_filtro_texto.delete(0, tk.END)
        self.combo_ext.current(0)                                                                #Vuelve a "Todos"
        self.actualizar_lista_visual(self.enlaces_encontrados)                                   #Mostrar lista completa
        self.lbl_estado_files.config(text=f"Total de Archivos: {len(self.enlaces_encontrados)}", fg="green")

    def descargar_archivos_seleccionados(self):                                                  
        for i, var in enumerate(self.check_vars):                                                
            if var.get():
                self.urls_seleccionadas.add(self.enlaces_visibles[i][1])
        
        indices_seleccionados = [i for i, var in enumerate(self.check_vars) if var.get()]        #Obtener los archivos seleccionados
        limite_descargas_manuales = 7                                                            #Arbitrario: Num. de archivos máximos por los cuales preguntar si se quieren nombrar o no.
        
        cantidad = len(indices_seleccionados)                                                    #Se cambia el if para trabajar con la variable cantidad
        if cantidad == 0:
            messagebox.showwarning("Cuidado", "No seleccionaste ningún archivo")
            return
        organizar_por_anio = self.var_organizar_anio.get()                                       #Para verificar si se va a organizar por carpetas

        #Apartado para descarga masiva manual o automática
        modo_descarga = "MANUAL"                                   #Por defecto.
        if organizar_por_anio:                                     #Si se descarga por ficheros / bibliotecas, se descarga de forma Automática
            modo_descarga = "AUTOMATICO"
        else:
            if cantidad == 1:                                      #Si solo es uno evita obviedades y manda a guardar pidiendo el nombre
                modo_descarga  = "MANUAL"
            else:
                if cantidad > limite_descargas_manuales:           #lim_des_man es un valor arbitrario, puede cambiarse
                    modo_descarga = "AUTOMATICO"                   #Si supera los 7, ya no preguntará si quiere nombrar a los archivos y se hará de forma automática 
                else:                                              #De 2 a 7 archivos entra a este caso  
                    respuesta =  messagebox.askyesno("Descarga de Archivos", f"Has seleccionado {cantidad} archivos. \n\n ¿Deseas nombrarlos uno por uno?\n Si eliges NO, se descargarán automáticamente con su nombre original.")
                    if not respuesta: #Si dice NO
                        modo_descarga = "AUTOMATICO"               #Igual da la opción de automatizarlo en el intervalo de 2 a 7 archivos

        carpeta_destino = ""
        if modo_descarga == "AUTOMATICO":                          #Verificación de que se descargarán N archivos
            if not organizar_por_anio:                             #Solo entra si no es descarga por ficheros y es mayor al limite nombrable            
                if cantidad > limite_descargas_manuales:
                    respuesta =  messagebox.askyesno("Descarga de Archivos", f"Has seleccionado {cantidad} archivos. \n\n ¿Deseas descargarlos a todos?\nPuede ser un proceso tardado.")
                    if not respuesta:                              #Si se arrepiente y cancela
                        return
            
            carpeta_destino = filedialog.askdirectory(title = "Selecciona la carpeta donde quieres descargar tus archivos.")        #Elección directorio
            if not carpeta_destino: return

        exito  = 0
        errores = 0
        self.lbl_estado_files.config(text="Descargando. . .", fg ="blue")
        self.root.update()

        session = requests.Session()
        for cookie in self.mis_cookies:
            session.cookies.set(cookie['name'], cookie['value'])

        for i in indices_seleccionados:                                                          #Iteramos sobre los seleccionados
            if i < len(self.enlaces_visibles):
                nombre_display, url = self.enlaces_visibles[i]
            else:
                continue    #Por si las mosquis
            
            try:
                nombre_sugerido = unquote(url.split('/')[-1])                                    #Sugerencia de nombre
                if "." not in nombre_sugerido:
                    nombre_sugerido += ".csv"  

                ruta_completa = ""
                # Creación de carpetas por año
                if organizar_por_anio:                                                                                      #Descarga si la checkbox está marcada

                    anios_encontrados = re.findall(r'\b(19\d{2}|20\d{2})\b', nombre_display + " " + nombre_sugerido)        #Formato de búsqueda de fechas
                    carpeta_anio = "Otros_Sin_Fecha"                                                                        #Archivos sin fecha encontrada 
                    carpeta_anio = anios_encontrados[0]                                                                     #Si encuentra dos fechas (2020-2025) va a tomar la primera 
                    ruta_directorio_anio = os.path.join(carpeta_destino, carpeta_anio)
                    os.makedirs(ruta_directorio_anio, exist_ok=True)                                                        #Crea la carpeta si no existe todavía

                    ruta_completa = os.path.join(ruta_directorio_anio, nombre_sugerido)
                
                elif modo_descarga == "MANUAL":
                    _, extension_original = os.path.splitext(nombre_sugerido)                        #Se consigue la extesnión del archivo
                    ext_limpia = extension_original.lower()
                    
                    opciones_guardado = []                                                           #Se da la opción de elegir la extensión, pero no cualquiera
                    if ext_limpia == ".pdf":                                                         #Da opciones relacionadas/parecidas a su extensión original
                        opciones_guardado = [("Documento PDF", "*.pdf")]                             #Para evitar errores de tipo de archivo
                    elif ext_limpia in [".zip", ".rar"]:
                        opciones_guardado = [("Archivo ZIP", "*.zip"),("Archivo RAR", "*.rar")]
                    elif ext_limpia in [".csv", ".xls", ".xlsx"]:
                        opciones_guardado = [("Excel", "*.xlsx"), ("CSV", "*.csv")]
                    else:
                        opciones_guardado = [("Todos los archivos", "*.*")]                          #Por cualquier cosa de deja la opción de Todos los archivos

                    ruta_completa = filedialog.asksaveasfilename(                                    #Ahora se abre una ventana de Guardar como para cada archivo
                    title=f"Guardar archivo: {nombre_display}...",                                   #PRUEBAAA [:50]
                    initialfile=nombre_sugerido,                                                     #Se sugiere el nombre original (esto porque hay archivos que se guardan de formas raras)
                    filetypes = opciones_guardado
                    )

                    if not ruta_completa:                                                            # Si el usuario da "Cancelar" en la ventana de guardar, ruta_completa estará vacía
                        continue                                                                     #Saltamos este archivo y seguimos con el siguiente (si hay más)

                    if not os.path.splitext(ruta_completa)[1]:              #Si el usuario cambia el nombre del archivo descargado y no se pone la extensión, se asigna la capturada anteriormente 
                        ruta_completa += extension_original                                          #Así se evita que tenga error de tipo de archivo
                else:                                                
                    ruta_completa = os.path.join(carpeta_destino, nombre_sugerido)                   #Descarga MODO AUTOMATICO pero sin carpetas 

                response  = session.get(url, stream=True)                                            #Si no jala con headers lo quito

                if response.status_code == 200:
                    ruta_temp = ruta_completa                                                       #Se guarda en una ruta temporal
                    conversion_necesaria = False                                                    #CONVERSIÓN DE XLS A CSV
                    es_destino_csv = ruta_completa.lower().endswith(".csv")
                    es_origen_excel  = ".xlsx" in extension_original.lower()
                    if es_destino_csv and es_origen_excel:            #Se compara la extensión y se decide si es necesario convertir
                        ruta_temp = ruta_completa + "_temp_excel"
                        conversion_necesaria  = True

                    with open(ruta_completa, 'wb') as f:
                        for chunk in response.iter_content(chunk_size=1024):
                            f.write(chunk)
                    
                    if conversion_necesaria:                                                        #Proceso de conversión
                        try:
                            df = pd.read_excel(ruta_temp)
                            df.to_csv(ruta_completa, index = False, encoding= 'utf-8-sig')
                            os.remove(ruta_temp)
                        except Exception as e:
                            print(f"Error de conversión a CSV.")                                    #Si falla, igual lo descargamos para no perder datos
                            if os.path.exists(ruta_temp):           
                                os.replace(ruta_temp, ruta_completa)

                    exito +=1
                else:
                    errores += 1
            except:
                errores += 1

        if exito > 0 or errores > 0:
            messagebox.showinfo("Resultados", f"Descargas exitosas: {exito}.\nErrores: {errores}.")
        else:
            messagebox.showinfo("Resultados", "No se guardó ningún archivo.")                           #Esto es cuando se cancelaron todas las ventanas de guardar

        self.lbl_estado_files.config(text="Proceso finalizado", fg="green")


    #Esta función es para el apartado de descargar tablas (Calidad del aire)
    def guardar_excel_csv(self, tablas, origen):
        archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
        if not archivo: return
        try:
            if archivo.endswith(".xlsx"):                                                               # t = tablas, c = columnas
                with pd.ExcelWriter(archivo) as writer:                                                 #exclewriter es para que cada tabla descargada sea una hoja de excel
                    for i, t in enumerate(tablas):                                                      #Ciclo for para iterar en cada tabla

                        if isinstance(t.columns, pd.MultiIndex):                                        #Aplanar MultiIndex, a cada tabla se le da el nombre 'Tabla i'
                            t.columns = [' '.join(map(str, c)).strip() for c in t.columns.values]
                        t.to_excel(writer, sheet_name=f"Tabla_{i+1}", index=False)                      #Guardamos cada tabla en su hoja correspondiente

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

    def reiniciar_interfaz(self):                                                                       #Esto es para que no se quede con los botones bloqueados y siga funcionando
        """Devuelve los botones a su estado original para hacer otra descarga"""
        self.btn_descargar_web.config(state="disabled")
        self.lbl_estado_web.config(text="Listo para nueva descarga.", fg="green")
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