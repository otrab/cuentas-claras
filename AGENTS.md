# Contexto: Cuentas Claras

Parseo de excel del banco + actualización de planilla en Google Drive para
registro y cálculo de arreglo de cuentas con mi esposa.

## Reglas
- Nunca modificar ni sobrescribir la planilla de gdrive sin mostrarme antes
  qué cambios se van a hacer
- Nunca inventar montos ni movimientos — si un dato del excel no es claro,
  preguntar en vez de asumir
- Antes de calcular un "arreglo de cuentas" (quién le debe a quién),
  mostrar el desglose completo, no solo el resultado final
- Datos financieros sensibles: no compartir montos fuera de este flujo
  sin que yo lo pida explícitamente

## Entorno
- Python 3.13 vía miniconda
- Librerías probables: pandas, openpyxl, gspread o google-api-python-client
  (instalar con pip3 si faltan)
