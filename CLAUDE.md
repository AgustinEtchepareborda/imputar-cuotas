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
| BOLSA CEMENTO | fila 3 | fila 4 | M (13) | I (9) |

**Columnas de pago**: auto-detectadas por el script buscando "MES AÑO teorico" en los headers. Las 4 columnas (teorico, real, N° cuota, fecha) son siempre consecutivas a partir de ahí.

Hoja USD (`$  USD fijo`, dos espacios): misma estructura, header en fila 3.

CUITs (en cualquier hoja) pueden tener formato con guiones (`20-11111111-2`), múltiples separados por `/`, `y` o texto libre (`27-11111111-2 (nombre) apellido 20-33333333-4`). El script extrae todos los CUITs de la celda automáticamente — un mismo CUIT puede aparecer en varios lotes (incluso con nombres distintos) y cada transferencia va a un lote libre.

## Cómo determina el script el número de cuota (regla actual)

**NO usar la columna A (MAYOR CUOTA)**. Esa fórmula no funciona correctamente para filas nuevas o cuando se imputan varias cuotas en la misma celda.

En cambio, el script:
1. Lee la columna "NUMERO DE CUOTA" del **mes actual** en la fila del cliente:
   - Si dice `"parte de cX"` → la próxima cuota entera es `X+1` (en pesos, un pago chico puede ir a completar X: ver "Partes de cuota")
   - Si es un número normal → ya está imputado este mes, no volver a imputar
2. Si está vacía, escanea **todas las columnas históricas** "NUMERO DE CUOTA" (meses anteriores) en la misma fila y toma el máximo + 1
3. Si no hay historial en deudores, busca en las últimas hojas de imputaciones (semanas previas) la cuota imputada para ese CUIT
4. Si aún no hay dato, usa `CUOTA_OVERRIDE` (dict manual en el CONFIG del script)

## Cómo resuelve el script clientes no encontrados por CUIT

Muchos clientes transfieren desde el número de cuenta (`402...`, `440...`), no desde su CUIT. Si el CUIT/cuenta de una transferencia no está en deudores:
1. Busca ese número en las **últimas 20 hojas** de imputaciones (6 si es USD) **y en las filas ya imputadas —amarillas— de la hoja actual** (las hojas USD son acumulativas: el antecedente del mismo cliente suele estar más abajo en la misma hoja) → extrae el nombre de col H
2. Del texto de col H se corta todo lo que va desde el primer marcador `cN`/`lN` en adelante (`"Fulano c8 en ambos lotes mismo monto"` → `"Fulano"`), porque las notas sueltas rompían la búsqueda
3. Busca ese nombre en deudores por palabras clave (case-insensitive)
4. Si hay 1 match → imputa; si hay **varios pero todos del mismo titular** → los trata como lotes del mismo cliente (sigue por la lógica de reparto); si hay 0 o varios con nombres distintos → reporta como ambiguo
5. Dos transferencias de **cuentas distintas** que resuelven al mismo cliente por nombre se imputan como cuotas consecutivas en la misma fila (la identidad es el nombre, no el número de cuenta)

## Cómo imputa múltiples lotes del mismo cliente

Si un CUIT tiene N lotes en deudores y llegan N transferencias iguales:
- Si todos los lotes tienen el **mismo nombre**: agrega `l{LOTE}` en el label (ej: `Apellido Nombre l6 c13`)
- Si los lotes tienen **nombres distintos**: usa el nombre del lote, sin aclarar lote
- Asigna lotes en orden (por hoja+fila) sin reutilizar el mismo lote en la misma corrida

**Una transferencia que cubre varios lotes** (ej: paga 2 cuotas de una): si el monto ≈ la suma de los teóricos de **N** lotes libres, reparte **una cuota a cada lote** (cada uno con su propio teórico y su propia numeración de cuota) y escribe `Nombre l13 c33 y l14 c23`. Prueba N de mayor a menor: no hace falta que el monto cubra *todos* los lotes del CUIT — un cliente puede figurar como cofirmante en un lote de otro titular que no se paga hace meses. Para elegir cuáles, prioriza los lotes con el pago más reciente (la columna "NUMERO DE CUOTA" más a la derecha con valor). Solo si ningún reparto cierra apila las N cuotas en un único lote (el de teórico más parecido).

## Estructura de imputaciones.xlsx

Cada hoja semanal ("S 121", "USD 5", etc.) tiene:
- Col A: Fecha
- Col F: Monto transferido (Crédito)
- Col C: Concepto — contiene el CUIT como número de 11 dígitos (o 12 para cuentas tipo `402XXXXXXXXX`)
- Col H: Nombre del cliente + cuota (se escribe acá al imputar)
- Col I: "x" (marca de procesado, se escribe acá al imputar)
- Fondo amarillo en la fila = ya procesado

## Reglas de negocio

- Solo procesar filas NO amarillas y sin "PAGO MENOS" en col H (salvo `reprocesar_pago_menos`)
- Identificar cliente por CUIT/CUIL extraído del campo Concepto
- Si no hay CUIT extraíble → reportar como ambiguo
- **Pesos**: si pagó MENOS y diferencia > tolerancia ($3.000) → "parte de cN" (ver "Partes de cuota")
- **USD**: si pagó MENOS y diferencia > U$D 5 → escribir "PAGO MENOS" en col H
- Si pagó menos pero diferencia ≤ tolerancia → imputar normalmente (cuota entera)
- Si pagó más → imputar normalmente
- Si ya fue imputado este mes → reportar como ambiguo (no sobreescribir)
- Al imputar: escribir en col H `{nombre} [l{lote}] c{cuota}`, en col I `x`, pintar fila amarilla

## BOLSA CEMENTO: teórico por tramo de fecha

El precio de la bolsa se congela del 1 al 10 de cada mes; después sube (sept-26: 1–10 $8.300, 11–15 $8.390, 16+ $8.490). La oficina cambiaba a mano la fórmula del teórico de cada fila según el día de pago. El script calcula el teórico él mismo:

`teórico = BOLSAS POR MES (col U) × precio del tramo del día de la transferencia × 2 si es bolsa de 50 kg`

- Tramos: encabezados de la **fila 1** del mes (`... 1 AL 10 DE SEPT 2026`, `... 10 al 15 ...`, `... desde el 16 ...`), precio en la **fila 2**. Se ignoran las columnas `los q firmaron hasta agosto 2025` (son el mismo precio ×2). Un `DESC <MES> 26` en fila 1 aplica como coeficiente. En la app se pueden sobreescribir (`10: 8300, 15: 8390, 31: 8490`).
- El tramo congelado se estira hasta el **día 13** (`DIA_TOPE_CONGELADO`): una transferencia del 10 que cae en finde/feriado aparece acreditada el 12 o 13. El mensaje de reclamo igual dice "después del día 10".
- Bolsa de 50 kg (×2): si la fórmula del teórico de la fila apunta a una columna "firmaron hasta agosto 2025"; si no se puede, fecha firma (col K) ≤ 31/08/2025.
- Otros meses / otras hojas: teórico de la planilla. Si el archivo no trae valores cacheados (lo guardó openpyxl), `LectorDeudores` evalúa las fórmulas aritméticas simples (`=U7*EX$2`, `=S7/T7`, `=497000+6000`).

## Partes de cuota (solo pesos)

Formato igual al que usa la oficina a mano. Col H: `Nombre parte de c33`, `Nombre completa c32`, `Nombre completa c5 y c6`, `Nombre c14 y completa c13`. Deudores: el "pago real" se va sumando como fórmula (`=330000+9600`), N° cuota `parte de c33` → al completarse `33` (o `32 y parte de c33`), la fecha queda la del **primer** pago (sirve para calcular el saldo al precio de ese día).

Una parte pendiente = celda "N° cuota" con `parte de cN` en los últimos 3 meses (`MESES_PARTE`) con saldo > tolerancia. Saldo = (cuotas que menciona la celda × teórico de ese mes) − pagado. Si se completa en un mes posterior, se suma en la columna del mes **de la parte**.

Criterio (acordado con el usuario): *si el monto parece la cuota del mes, es la cuota del mes; si no, se cancela primero lo más viejo.*

Qué hace con cada transferencia (un lote destino):
1. **Cuota entera** (±tolerancia) o múltiplo → cuota(s) siguiente(s) `X+1`, **sin mirar partes pendientes** (X = mayor cuota mencionada, incluidas las parciales).
2. **Paga de menos**:
   - sin partes pendientes, o el pago es ≥ 90% del teórico (`UMBRAL_CUOTA_DEL_MES`) y supera lo adeudado → `parte de cX+1` (la cuota del mes pagada de menos, típico: precio congelado pagado después del 10).
   - si no → completa partes pendientes (la más vieja primero); si no alcanza, acumula en esa parte; lo que sobra queda `parte de cX+1`.
3. **Paga de más**: si el sobrante alcanza para completar partes pendientes → `cX+1 y completa cN`; si no, como siempre (normal si ≤ $50.000, si no "PAGO MAS").

**Reclamo por WhatsApp** (solo BOLSA CEMENTO): cada cuota que queda incompleta genera un mensaje con el saldo (y las otras partes abiertas) y un link `wa.me` al teléfono de la col TELEFONO (normalizado a `549…`; `telefonos_extra` permite pasarle teléfonos de otra fuente, p. ej. Supabase).

## Para cambiar de mes

No hace falta hacer nada — el script detecta el mes a partir de las fechas de las transferencias y busca la columna correspondiente en deudores automáticamente. Si el archivo de deudores no tiene ese mes aún, el script lanza un error explicativo.
