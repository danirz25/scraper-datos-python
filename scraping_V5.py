#Esta versión además de la interfaz gráfica, trabaja con PDF's
import pandas as pd
import os
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog
import pdfplumber                   #Para trabajar con PDF
from io import  StringIO            #Para tratar cadenas de texto como archivos
from selenium import webdriver      #Selenium es para poder trabajar con JavaScripts en la web
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options


class ScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Descargador de Tablas Web - PDF")
        self.root.geometry("600x450")
        self.driver = None

        self.tab_control = ttk.Notebook(root)            #Para que se distribuya mejor la infterfaz, se va a dividir en pestañas

        self.tab_web = tk.Frame(self.tab_control)       #PESTAÑA 1: Para sitios web
        self.tab_control.add(self.tab_web, text= ' Sitios Web ')

        self.tab_pdf = tk.Frame(self.tab_control)       #PESTAÑA 2: Para archivos PDF de forma local
        self.tab_control.add(self.tab_pdf, text= ' Archivos PDF ')

        self.tab_control.pack(expand=1, fill="both")
        
        #Pestaña Web
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

        #Pestaña PDF
        self.lbl_inst_pdf = tk.Label(self.tab_pdf, text = "Selecciona el archivo PDF desde tu computadora: ", font=("Arial", 12))
        self.lbl_inst_pdf.pack(pady=(30,10))

        self.btn_cargar_pdf = tk.Button(self.tab_pdf, text="Seleccionar", command = self.procesar_pdf, bg="#FF9800", fg="white", height=2, width=30)
        self.btn_cargar_pdf.pack(pady=10)

        self.lbl_estado_pdf = tk.Label(self.tab_pdf, text="No se eligió ningún archivo", fg="gray")
        self.lbl_estado_pdf.pack(pady = 20)

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

            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])  #Esto es algo anti bots para intentar entrar a links más protegidos
            chrome_options.add_experimental_option('useAutomationExtension', False)
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")

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
            html = self.driver.page_source            #Obtenemos el html completo
            self.driver.quit()                        #Se cierra el navegador 
            self.driver = None                        #Se optimizó esta ejec de lineas 97-101
            
            tablas = pd.read_html(StringIO(html))       # Le pasamos el TEXTO del html a pandas (en vez de la url directa, no sé por qué pero con url dejó de funcionar xd)

            if not tablas:                              #En lugar de comparar el num. de tablas ahora se hace directo xd 
                messagebox.showwarning("Aviso", "No se encontraron tablas en la página actual.")   #mensaje por si no encuentra tablas 
                self.reiniciar_interfaz()
                return

            self.guardar_datos(tablas, "WEB")       #Se hizo una función mejor para guardar tanto en web como html
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

    def procesar_pdf(self):                 #Función para procesar archivos PDF
        archivo_pdf = filedialog.askopenfilename (
            title = "Seleccionar Archivo PDF.",
            filetypes=[("Archivos PDF", "*.pdf")]
            )
        if not archivo_pdf:
            return
        self.lbl_estado_pdf.config(text=f"Procesando: {os.path.basename(archivo_pdf)}...", fg="blue")
        self.root.update()
    
        try:
            tablas_encontradas = []
            with pdfplumber.open(archivo_pdf) as pdf:       #Se abre el pdf
                for i, pagina in enumerate(pdf.pages):
                    tablas_pag = pagina.extract_tables()   #Extrae tablas de la página actual

                    for tabla_raw in tablas_pag:
                        df = pd.DataFrame(tabla_raw[1:], columns=tabla_raw[0])      #Convertir a Dataframe 
                        df = df.dropna(how= 'all')                                  #Se asume que la primer fila son los encabezados
                    
                        if not df.empty:
                            tablas_encontradas.append(df)                           #Elimina filas totalmente vacías

            if not tablas_encontradas:                                              #Si no hay tablas, ps vale queso
                messagebox.showwarning("PDF","No se encontraron tablas en el archivo")
                self.lbl_estado_pdf.config(text="Ninguna tabla encontrada.", fg = "red")
                return
        
            self.guardar_datos(tablas_encontradas, "PDF")                           #Si hay tablas, las guarda.
            self.lbl_estado_pdf.config(text="Proceso realizado con éxito.", fg="green")
        except Exception as e:
            messagebox.showerror("Error," f"No se pudo leer el PDF: \n{e}")
            self.lbl_estado_pdf.config(text="Error", fg="red")
            
    def guardar_datos(self, tablas, origen):
        archivo_guardar = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Archivos CSV", "*.csv")],
            title=f"Guardar tablas extraidas de {origen}"
        )
        if not archivo_guardar: return

        guardados = 0
        mensaje_final = ""

        try:
            #Para guardar el Excel
            if archivo_guardar.endswith(".xlsx"):
                with pd.ExcelWriter(archivo_guardar) as writer:         #exclewriter es para que cada tabla descargada sea una hoja de excel 
                    for i, tabla in enumerate(tablas):                  #Ciclo for para iterar en cada tabla
                      nombre_hoja = f"Tabla_{i+1}"                    #A cada tabla se le da el nombre 'Tabla i'
                      if isinstance(tabla.columns, pd.MultiIndex):    #se añade esto para 'aplanar' las celdas combinada porque si no, da error.
                         tabla.columns = [' '.join(map(str, col)).strip() for col in tabla.columns.values]
                    
                      if not tabla.empty:                             #El if es para evitar guardar tablas vacias 
                         tabla.to_excel(writer, sheet_name=nombre_hoja, index=False)   # Guardamos cada tabla en su hoja correspondiente
                         guardados += 1
                mensaje=f"Se guardó el excel corretamente con {guardados} tablas. "
            
            #Para guardar CSV
            elif archivo_guardar.endswith(".csv"):
                base_path = os.path.splitext(archivo_guardar)[0]
                for i, tabla in enumerate(tablas):
                    if isinstance(tabla.columns, pd.MultiIndex):
                        tabla.columns = [' '.join(map(str, col)).strip() for col in tabla.columns.values]

                    if not tabla.empty:
                        nombre = f"{base_path}_Tabla{i+1}.csv" if len(tablas) > 1 else archivo_guardar
                        tabla.to_csv(nombre, index = False, encoding= 'utf-8-sig')
                        guardados += 1
                mensaje = f"Se genreraron {guardados} archivos CSV."
            messagebox.showinfo("Éxito", f"{mensaje}\nOrigen: {origen}")
        except Exception as e:
            messagebox.showerror("Error: Algo salió mal", str(e))

#Bloque de ejecución
if __name__ == "__main__":
    ventana = tk.Tk()
    app = ScraperApp(ventana)
    ventana.mainloop()


#Para Comentar varias lienas es Ctrl + K + C 
#Para Descomentar varias lineas es Ctrl + K + U 

