from tkinter import *
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment
import xml.etree.ElementTree as ET
import pandas as pd
import unidecode
import unicodedata
import psycopg2
import uuid
from psycopg2 import Error

connection = psycopg2.connect(database="banco2",
                                        host="localhost",
                                        user="postgres",
                                        password="Gumattos2",
                                        port="5432")
cursor = connection.cursor()

cursor.execute("SELECT COUNT(*) FROM congressos WHERE id_professor = 'e9a2f9cf-0edb-4f2a-9879-e1c43cf6ab10' AND anoconclusao = '2003' ")
artigos = cursor.fetchone()[0]

print(artigos)