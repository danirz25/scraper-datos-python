import bs4
import pandas as pd 
import numpy as np 
import matplotlib.pyplot as plt 
import seaborn as sns

from urllib.request import urlopen
from bs4 import BeautifulSoup

#VERSION 1. poner la url manual
#Link de referencia http://www.hubertiming.com/results/2017GPTR10K
url = "https://calidaddelaire.puebla.gob.mx/views/reporteICA.php"   #Link donde queremos descargar la tabla/informacion
html = urlopen(url)

soup= BeautifulSoup(html, 'lxml')
type(soup)
bs4.BeautifulSoup

#Sacar titulo
title = soup.title
print(title)

#Sacar links 
#all_links = soup.find_all("a")
#for link in all_links:
#    print(link.get("href"))

# Imprimir N filas
rows = soup.find_all('tr')
print(rows[:50]) #N

#VERSION 2: poner la url en terminal

#VERSION 3: poner la url en interfaz