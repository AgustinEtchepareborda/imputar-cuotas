# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Propósito del proyecto

Imputación automática de cuotas: tomar las transferencias bancarias del archivo `imputaciones.xlsx` (hoja semanal como "S 121", "S 122", etc. o "USD 5", "USD 6"...) y registrarlas en `deudores.xlsx` (monto real, número de cuota, fecha).

## Archivos clave

```
datos/
  imputaciones.xlsx            ← archivo de transferencias bancarias
  deudores.xlsx                ← planilla de deudores/cuotas
  comprobantes_cache.json      ← cache local del Google Sheet "Datos comprobantes"
imputar_s120.py                ← script pesos (listo para S 121, DRY_RUN=True)
imputar_usd5.py                ← script USD (listo para USD 5, DRY_RUN=True)
comprobantes_helper.py         ← módulo de acceso al cache de comprobantes
exportar_comprobantes.py       ← refresca el cache desde Google Sheets
```

## Cache de comprobantes (Google Sheets)

El sheet "Datos comprobantes" (su ID se configura en la variable de entorno `COMPROBANTES_SHEET_ID`) tiene los datos de clientes que mandaron comprobantes al bot. Se usa como fallback adicional cuando el CUIT no se encuentra en deudores ni en hojas anteriores de imputaciones.

**Antes de cada imputación semanal**, refrescar el cache:
```
python exportar_comprobantes.py
```
Requiere `pip install gspread`. La primera vez abre el navegador para autorizar (OAuth). Después queda guardado.

Si no se corre el exportar, igual funciona con el último cache guardado en `datos/comprobantes_cache.json`.

## Cómo correr la imputación de una nueva semana (pesos)

El script **auto-detecta** la hoja (última "S NNN" del archivo) y las columnas del mes (busca "MES AÑO teorico" en los headers de deudores). No hace falta tocar nada al cambiar de semana o de mes.

1. **Solo revisar `DRY_RUN = True`** en `imputar_s120.py` (ya está así por defecto)
2. **Correr dry-run**: `python imputar_s120.py`
3. **Revisar el reporte**, especialmente los casos ambiguos
4. **Cambiar `DRY_RUN = False`** y volver a correr para escribir

Si necesitás procesar una hoja específica (no la última), setear `IMP_SHEET = 'S 123'` manualmente.

## Cómo correr la imputación USD

Igual que pesos — auto-detecta la última hoja "USD N" y el mes:

1. `DRY_RUN = True` en `imputar_usd5.py`
2. `python imputar_usd5.py`
3. Revisar → `DRY_RUN = False` → volver a correr

## Estructura de deudores.xlsx

Hojas con cuotas en pesos (NO escribir en '$ USD fijo' desde el script pesos):

| Hoja | Header row | Data desde | CUIT col | Nombre col |
|------|-----------|------------|----------|-----------|
| INDICE CAC | fila 5 | fila 6 | L (12) | I (9) |
| INDICE CAC M. OBRA | fila 3 | fila 4 | K (11) | I (9) |
| BOLSA CEMENTO | fila 3 | fila 4 | N (14) | J (10) |

**Columnas de pago**: auto-detectadas por el script buscando "MES AÑO teorico" en los headers. Las 4 columnas (teorico, real, N° cuota, fecha) son siempre consecutivas a partir de ahí.

Hoja USD (`$  USD fijo`, dos espacios): misma estructura, header en fila 3.

CUITs (en cualquier hoja) pueden tener formato con guiones (`20-11111111-2`), múltiples separados por `/`, `y` o texto libre (`27-11111111-2 (nombre) apellido 20-33333333-4`). El script extrae todos los CUITs de la celda automáticamente — un mismo CUIT puede aparecer en varios lotes (incluso con nombres distintos) y cada transferencia va a un lote libre.

## Cómo determina el script el número de cuota (regla actual)

**NO usar la columna A (MAYOR CUOTA)**. Esa fórmula no funciona correctamente para filas nuevas o cuando se imputan varias cuotas en la misma celda.

En cambio, el script:
1. Lee la columna "NUMERO DE CUOTA" del **mes actual** en la fila del cliente:
   - Si dice `"parte de cuota X"` → imputa como `X+1` (el cliente pagó parte de X, ahora completa)
   - Si es un número normal → ya está imputado este mes, no volver a imputar
2. Si está vacía, escanea **todas las columnas históricas** "NUMERO DE CUOTA" (meses anteriores) en la misma fila y toma el máximo + 1
3. Si no hay historial en deudores, busca en las últimas hojas de imputaciones (semanas previas) la cuota imputada para ese CUIT
4. Si aún no hay dato, usa `CUOTA_OVERRIDE` (dict manual en el CONFIG del script)

## Cómo resuelve el script clientes no encontrados por CUIT

Si el CUIT de una transferencia no está en deudores:
1. Busca ese CUIT en las últimas 8 hojas de imputaciones → extrae el nombre de col H
2. Busca ese nombre en deudores por palabras clave (case-insensitive)
3. Si hay 1 match → imputa; si hay 0 o varios → reporta como ambiguo

## Cómo imputa múltiples lotes del mismo cliente

Si un CUIT tiene N lotes en deudores y llegan N transferencias iguales:
- Si todos los lotes tienen el **mismo nombre**: agrega `l{LOTE}` en el label (ej: `Apellido Nombre l6 c13`)
- Si los lotes tienen **nombres distintos**: usa el nombre del lote, sin aclarar lote
- Asigna lotes en orden (por hoja+fila) sin reutilizar el mismo lote en la misma corrida

## Estructura de imputaciones.xlsx

Cada hoja semanal ("S 121", "USD 5", etc.) tiene:
- Col A: Fecha
- Col F: Monto transferido (Crédito)
- Col C: Concepto — contiene el CUIT como número de 11 dígitos (o 12 para cuentas tipo `402XXXXXXXXX`)
- Col H: Nombre del cliente + cuota (se escribe acá al imputar)
- Col I: "x" (marca de procesado, se escribe acá al imputar)
- Fondo amarillo en la fila = ya procesado

## Reglas de negocio

- Solo procesar filas NO amarillas y sin "PAGO MENOS" en col H
- Identificar cliente por CUIT/CUIL extraído del campo Concepto
- Si no hay CUIT extraíble → reportar como ambiguo
- Si pagó MENOS y diferencia > $3.000 (pesos) o > U$D 5 (USD) → escribir "PAGO MENOS" en col H
- Si pagó menos pero diferencia ≤ tolerancia → imputar normalmente
- Si pagó más → imputar normalmente
- Si ya fue imputado este mes → reportar como ambiguo (no sobreescribir)
- Al imputar: escribir en col H `{nombre} [l{lote}] c{cuota}`, en col I `x`, pintar fila amarilla

## Para cambiar de mes

No hace falta hacer nada — el script detecta el mes a partir de las fechas de las transferencias y busca la columna correspondiente en deudores automáticamente. Si el archivo de deudores no tiene ese mes aún, el script lanza un error explicativo.
