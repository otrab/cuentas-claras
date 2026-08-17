#!/usr/bin/env python3
"""Diagnóstico rápido: qué devuelve gspread para la columna B de 'resumen'."""
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials

CARPETA = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_CREDENCIALES = os.path.join(CARPETA, "credentials.json")
NOMBRE_SHEET = "cuentas claras"

scope = ["https://spreadsheets.google.com/feeds",
          "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(ARCHIVO_CREDENCIALES, scope)
gc = gspread.authorize(creds)
spreadsheet = gc.open(NOMBRE_SHEET)
ws = spreadsheet.worksheet("resumen")

valores_b = ws.col_values(2)
print(f"Total valores en columna B: {len(valores_b)}")
print("Últimos 6 valores (repr para ver el tipo/formato exacto):")
for v in valores_b[-6:]:
    print(f"  {v!r}  (tipo: {type(v).__name__})")

print()
print("Todos los nombres de hoja en el spreadsheet:")
print([ws.title for ws in spreadsheet.worksheets()])
