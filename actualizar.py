#!/usr/bin/env python3
"""
Actualizar Cuentas Claras
==========================
Lee todas las cartolas .xls en cartolas/, detecta qué meses todavía no
existen como hoja en el Google Sheet "cuentas claras", y para cada mes
nuevo arma los datos y los sube -- pero SIEMPRE mostrando el desglose
completo y pidiendo confirmación antes de escribir en el Sheet.

Uso:
    python3 actualizar.py

Requisitos en esta misma carpeta:
    - llaves.txt      (una descripción a excluir por línea)
    - credentials.json (credenciales de service account de Google)
    - cartolas/        (carpeta con los .xls descargados del banco)
"""

import os
import sys
import glob
import calendar
from datetime import datetime

import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

CARPETA = os.path.dirname(os.path.abspath(__file__))
CARPETA_CARTOLAS = os.path.join(CARPETA, "cartolas")
ARCHIVO_LLAVES = os.path.join(CARPETA, "llaves.txt")
ARCHIVO_CREDENCIALES = os.path.join(CARPETA, "credentials.json")
NOMBRE_SHEET = "cuentas claras"

MESES_ES = {
    1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Sept", 10: "Oct", 11: "Nov", 12: "Dic",
}

FORMATO_MONEDA_CARGO = '"$"#,##0'
FORMATO_MONEDA_INGRESO = '[$$]#,##0'
FORMATO_FECHA = "dd-mm"


# --------------------------------------------------------------------------
# Lectura y parseo de cartolas
# --------------------------------------------------------------------------

def leer_llaves():
    if not os.path.exists(ARCHIVO_LLAVES):
        print(f"⚠️  No encontré {ARCHIVO_LLAVES}. Sigo sin filtro de llaves.")
        return []
    with open(ARCHIVO_LLAVES, "r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]


def parsear_cartola(path_xls):
    """Devuelve una lista de dicts: {fecha (datetime), descripcion, monto, tipo}
    tipo es 'cargo' o 'abono'. Usa la fecha de emisión real dentro del archivo
    para saber a qué año/mes(es) pertenecen las filas (la cartola solo trae
    dd/mm, sin año)."""
    raw = pd.read_excel(path_xls, header=None)

    # Fecha de emisión: fila 13 (0-indexed), columna 4 -> "31/07/2026"
    fecha_emision_str = str(raw.iloc[13, 4]).strip()
    fecha_emision = datetime.strptime(fecha_emision_str, "%d/%m/%Y")
    anio_emision = fecha_emision.year
    mes_emision = fecha_emision.month

    # Tabla de movimientos empieza en fila 24 (0-indexed): headers
    tabla = raw.iloc[24:].copy()
    tabla.columns = raw.iloc[24]
    tabla = tabla.iloc[1:]
    tabla = tabla.dropna(subset=["Fecha"])

    movimientos = []
    for _, row in tabla.iterrows():
        desc = str(row["Descripción"]).strip()
        if desc in ("SALDO INICIAL", "SALDO FINAL", "nan"):
            continue

        fecha_str = str(row["Fecha"]).strip()  # "dd/mm"
        try:
            dia, mes = [int(x) for x in fecha_str.split("/")]
        except ValueError:
            continue

        # Si el mes de la fila es distinto (mayor) al mes de emisión,
        # significa que la cartola cruza fin de año hacia atrás.
        anio = anio_emision
        if mes > mes_emision:
            anio = anio_emision - 1
        fecha = datetime(anio, mes, dia)

        cargo = row.get("Cargos (PESOS)")
        abono = row.get("Abonos (PESOS)")

        if pd.notna(cargo):
            movimientos.append({
                "fecha": fecha, "descripcion": desc,
                "monto": float(cargo), "tipo": "cargo",
            })
        elif pd.notna(abono):
            movimientos.append({
                "fecha": fecha, "descripcion": desc,
                "monto": float(abono), "tipo": "abono",
            })

    return movimientos


def filtrar_llaves(movimientos, llaves):
    if not llaves:
        return movimientos
    return [m for m in movimientos if m["descripcion"] not in llaves
            and not any(llave in m["descripcion"] for llave in llaves)]


def agrupar_por_mes(movimientos):
    """Devuelve dict {(anio, mes): [movimientos]}"""
    grupos = {}
    for m in movimientos:
        clave = (m["fecha"].year, m["fecha"].month)
        grupos.setdefault(clave, []).append(m)
    return grupos


def nombre_hoja(anio, mes):
    return f"{MESES_ES[mes]} {anio}"


# --------------------------------------------------------------------------
# Google Sheets
# --------------------------------------------------------------------------

def conectar_sheet():
    if not os.path.exists(ARCHIVO_CREDENCIALES):
        sys.exit(f"❌ No encontré {ARCHIVO_CREDENCIALES}. Necesito las "
                  f"credenciales de la service account para conectar a Google Sheets.")
    scope = ["https://spreadsheets.google.com/feeds",
              "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(ARCHIVO_CREDENCIALES, scope)
    gc = gspread.authorize(creds)
    return gc.open(NOMBRE_SHEET)


def hojas_de_mes_existentes(spreadsheet):
    """Devuelve el set de nombres de hoja que parecen ser hojas de mes
    (para no confundir con 'resumen' o 'auto y mac')."""
    existentes = set()
    for ws in spreadsheet.worksheets():
        existentes.add(ws.title)
    return existentes


# --------------------------------------------------------------------------
# Construcción de la hoja del mes
# --------------------------------------------------------------------------

def crear_hoja_mes(spreadsheet, anio, mes, movimientos):
    titulo = nombre_hoja(anio, mes)
    n_filas = len(movimientos)
    ws = spreadsheet.add_worksheet(title=titulo, rows=str(max(n_filas + 10, 100)), cols="12")

    ultima_fila = n_filas + 2  # datos empiezan en fila 3

    # Fila 1: totales
    ws.update("B1:G1", [[
        "Gasto total", f"=sum(C2:C{ultima_fila})", f"=sum(D2:D{ultima_fila})",
        "Ingreso total", f"=sum(F2:F{ultima_fila})", f"=sum(G2:G{ultima_fila})",
    ]])

    # Fila 2: headers
    ws.update("A2:G2", [[
        "Fecha", "Descripción", "Cargos Seba", "Cargos Pía",
        "", "Ingreso Seba", "Ingreso Pía",
    ]])

    # Filas de datos (ordenadas por fecha)
    movimientos_ordenados = sorted(movimientos, key=lambda m: m["fecha"])
    filas = []
    for m in movimientos_ordenados:
        fecha_str = m["fecha"].strftime("%Y-%m-%d")
        if m["tipo"] == "cargo":
            filas.append([fecha_str, m["descripcion"], m["monto"], "", "", "", ""])
        else:
            filas.append([fecha_str, m["descripcion"], "", "", "", m["monto"], ""])

    if filas:
        ws.update(f"A3:G{2 + len(filas)}", filas, value_input_option="USER_ENTERED")

    # Formato de moneda y fecha
    ws.format(f"C1:D{ultima_fila}", {"numberFormat": {"type": "CURRENCY", "pattern": FORMATO_MONEDA_CARGO}})
    ws.format(f"F1:G{ultima_fila}", {"numberFormat": {"type": "CURRENCY", "pattern": FORMATO_MONEDA_INGRESO}})
    ws.format(f"A3:A{ultima_fila}", {"numberFormat": {"type": "DATE", "pattern": FORMATO_FECHA}})

    print(f"   ✅ Hoja '{titulo}' creada con {n_filas} movimientos.")
    return ws


def asegurar_fila_resumen(spreadsheet, anio, mes):
    """Crea (si no existe) la fila del mes en 'resumen', enlazada a la hoja
    de ese mes vía fórmulas, siguiendo el patrón vigente (iferror)."""
    ws = spreadsheet.worksheet("resumen")
    titulo_mes = nombre_hoja(anio, mes)

    valores_b = ws.col_values(2)  # columna B = Mes (fechas)
    fecha_objetivo = datetime(anio, mes, 1)

    for i, val in enumerate(valores_b[2:], start=3):  # datos desde fila 3
        try:
            fecha_fila = datetime.strptime(val, "%Y-%m-%d") if isinstance(val, str) else None
        except ValueError:
            fecha_fila = None
        if fecha_fila and fecha_fila.year == anio and fecha_fila.month == mes:
            print(f"   ℹ️  La fila de {titulo_mes} ya existe en 'resumen' (fila {i}). No se toca.")
            return

    fila = len(valores_b) + 1

    fila_datos = [
        False,                              # A: Cuenta clara?
        fecha_objetivo.strftime("%Y-%m-%d"),  # B: Mes
        f"=SUM(D{fila}:E{fila})",            # C: Ingreso Familiar
        f"='{titulo_mes}'!F1",               # D: Ingreso Seba
        "",                                   # E: Ingreso Pía (a mano)
        f"=SUM(I{fila}:J{fila})",            # F: Gasto Familia
        f'=iferror(D{fila}/C{fila},"")*100', # G: %S
        f'=iferror(E{fila}/C{fila},"")*100', # H: %P
        f"='{titulo_mes}'!C1",               # I: Gasto Seba
        f"='{titulo_mes}'!D1",               # J: Gasto Pia
        f'=iferror((D{fila}/C{fila})*F{fila},"")',  # K: Pago Total Seba
        f'=iferror((E{fila}/C{fila})*F{fila},"")',  # L: Pago Total Pía
        f"=K{fila}-I{fila}",                 # M: AJUSTE / Diferencia Seba
        f"=L{fila}-J{fila}",                 # N: Diferencia Pía
        f"=C{fila}-F{fila}",                 # O: Capacidad de ahorro
        f"=D{fila}-K{fila}",                 # P: Ahorro S
        f'=iferror(E{fila}-L{fila},"")',     # Q: Ahorro P
    ]

    ws.update(f"A{fila}:Q{fila}", [fila_datos], value_input_option="USER_ENTERED")
    print(f"   ✅ Fila de '{titulo_mes}' agregada a 'resumen' (fila {fila}).")


def crear_fila_mes_siguiente_si_falta(spreadsheet, anio, mes):
    """Prepara por adelantado la fila del mes siguiente en 'resumen',
    tal como es la costumbre actual, en False/vacía."""
    if mes == 12:
        anio_sig, mes_sig = anio + 1, 1
    else:
        anio_sig, mes_sig = anio, mes + 1
    asegurar_fila_resumen(spreadsheet, anio_sig, mes_sig)


# --------------------------------------------------------------------------
# Flujo principal
# --------------------------------------------------------------------------

def mostrar_desglose(anio, mes, movimientos):
    titulo = nombre_hoja(anio, mes)
    total_cargos = sum(m["monto"] for m in movimientos if m["tipo"] == "cargo")
    total_abonos = sum(m["monto"] for m in movimientos if m["tipo"] == "abono")

    print(f"\n{'='*60}")
    print(f"  MES NUEVO DETECTADO: {titulo}  ({len(movimientos)} movimientos)")
    print(f"{'='*60}")
    for m in sorted(movimientos, key=lambda x: x["fecha"]):
        signo = "-" if m["tipo"] == "cargo" else "+"
        print(f"  {m['fecha'].strftime('%d-%m')}  {signo}${m['monto']:>12,.0f}  {m['descripcion']}")
    print(f"{'-'*60}")
    print(f"  Total cargos (gasto Seba):   ${total_cargos:,.0f}")
    print(f"  Total abonos (ingreso Seba): ${total_abonos:,.0f}")
    print(f"{'='*60}\n")


def main():
    llaves = leer_llaves()

    archivos = sorted(glob.glob(os.path.join(CARPETA_CARTOLAS, "*.xls")))
    if not archivos:
        sys.exit(f"❌ No hay archivos .xls en {CARPETA_CARTOLAS}")

    print(f"📂 Encontré {len(archivos)} cartola(s) en {CARPETA_CARTOLAS}")

    todos_los_movimientos = []
    for archivo in archivos:
        print(f"   Leyendo {os.path.basename(archivo)}...")
        movs = parsear_cartola(archivo)
        movs = filtrar_llaves(movs, llaves)
        todos_los_movimientos.extend(movs)

    grupos = agrupar_por_mes(todos_los_movimientos)

    print("\n🔗 Conectando a Google Sheets...")
    spreadsheet = conectar_sheet()
    existentes = hojas_de_mes_existentes(spreadsheet)

    meses_nuevos = {}
    for (anio, mes), movs in sorted(grupos.items()):
        titulo = nombre_hoja(anio, mes)
        if titulo in existentes:
            print(f"⏭️  {titulo} ya existe en el Sheet. Se omite (no se toca).")
        else:
            meses_nuevos[(anio, mes)] = movs

    if not meses_nuevos:
        print("\n✅ No hay meses nuevos por agregar. Todo al día.")
        return

    for (anio, mes), movs in sorted(meses_nuevos.items()):
        mostrar_desglose(anio, mes, movs)
        resp = input("¿Subir este mes al Google Sheet? (s/n): ").strip().lower()
        if resp != "s":
            print("   ⏭️  Omitido por decisión del usuario.\n")
            continue

        crear_hoja_mes(spreadsheet, anio, mes, movs)
        asegurar_fila_resumen(spreadsheet, anio, mes)
        crear_fila_mes_siguiente_si_falta(spreadsheet, anio, mes)

    print("\n🎉 Listo.")


if __name__ == "__main__":
    main()
