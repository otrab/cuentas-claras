# Cuentas Claras — actualizador automático

## Instalación (una sola vez)

```bash
cd cuentas_claras
pip3 install -r requirements.txt
```

Poné tu `credentials.json` (service account de Google con acceso al
Sheet "cuentas claras") en esta misma carpeta.

## Uso normal

1. Bajás la cartola del banco (.xls) y la tirás, sin renombrar nada,
   dentro de la carpeta `cartolas/`.
2. Corrés:
   ```bash
   python3 actualizar.py
   ```
3. El script:
   - Lee todas las cartolas que hay en `cartolas/`.
   - Detecta el mes/año real de cada movimiento usando la fecha de
     emisión que viene dentro del propio Excel (no depende del nombre
     del archivo).
   - Revisa qué meses ya existen como hoja en el Google Sheet.
   - Para cada mes que **no existe todavía**, te muestra en la consola
     el desglose completo (fecha, monto, descripción) y te pregunta
     `¿Subir este mes al Google Sheet? (s/n)` antes de escribir nada.
   - Si decís que sí, crea la hoja del mes (con el mismo formato que
     usan hoy: headers, fórmulas de suma, formato de moneda) y agrega
     la fila correspondiente en `resumen`, además de dejar preparada
     (en blanco) la fila del mes siguiente.
   - Los meses que ya tienen hoja se saltan siempre — nunca se
     sobrescriben.

## Sobre `llaves.txt`

Una descripción (o fragmento de descripción) por línea. Cualquier
movimiento que coincida se excluye del todo (no se sube como cargo
ni como abono). Editalo cuando quieras agregar o sacar una llave.

## Sobre las cartolas viejas

Podés dejar todas las cartolas históricas en `cartolas/` sin
problema — el script las vuelve a leer cada vez, pero como ya
verifica qué meses existen en el Sheet, simplemente las va a saltar.
Así también te queda el archivo histórico ordenado en un solo lugar.

## Notas importantes

- Este MVP asume que la cartola es solo de Sebastián: todo cargo va
  a "Cargos Seba" y todo abono a "Ingreso Seba". Pía sigue agregando
  sus propios movimientos a mano en Google Sheets.
- La columna "Cuenta clara?" (A) queda siempre en blanco/False al
  crear una fila nueva — la marcan ustedes a mano cuando cuadran el
  mes.
