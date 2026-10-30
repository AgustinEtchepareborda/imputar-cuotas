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
app.py                         ← interfaz web (Streamlit), lo que usa el usuario
imputar_core.py                ← lógica de imputación (pesos y USD)
mep_helper.py                  ← cotización MEP (USD pagados en pesos)
comprobantes_helper.py         ← módulo de acceso al cache de comprobantes
exportar_comprobantes.py       ← refresca el cache desde Google Sheets
```

## ⚠️ El repo es PÚBLICO: nada de datos reales

Todo lo que se commitea queda visible en GitHub (y en el historial aunque después se borre). **Nunca** subir:

- Planillas, exports o caches (`*.xlsx`, `*.csv`, `*.pdf`, `*.json` de datos, logs de la app): van en `datos/` o fuera del repo, ignorados por `.gitignore`.
- Datos de clientes **ni siquiera como ejemplo** en código, docstrings, tests, comentarios, mensajes de commit o este archivo: nombres, CUIT/DNI, teléfonos, mails, números de cuenta, lotes con nombre, montos de transferencias concretas.
- Credenciales, tokens, IDs de Google Sheets o Supabase: van en variables de entorno o `.streamlit/secrets.toml`.

Para ejemplos usar datos inventados y obvios: CUIT `20-11111111-2`, teléfono `2990001111`, nombre `Apellido Nombre`. Antes de commitear, revisar el diff buscando datos copiados de las planillas. Si algo se filtró: sacarlo, reescribir el historial y hacer force push (ya pasó con teléfonos en un docstring, oct-2026).

## Comprobantes del bot (Google Sheets)

El sheet "Datos comprobantes" (lo llenan los bots de WhatsApp y mail de Leuche) es el **último recurso** para saber quién es el cliente de una transferencia: después de deudores y de las semanas anteriores. Solo da el **nombre**; la fila sale de buscarlo en deudores.

**Acceso**: la app lo lee directo con una cuenta de servicio de Google (proyecto Leuche, solo **Lector** de ese sheet). La clave va en `.streamlit/secrets.toml` (local, ignorado por git) y en los secrets de Streamlit Cloud:
```
COMPROBANTES_SHEET_ID = "..."
[gcp_service_account]
type = "service_account"
...  # el JSON de la clave, campo por campo
```
Se relee cada 10 min. Si no se puede leer, usa `datos/comprobantes_cache.json` (`python exportar_comprobantes.py` lo refresca). Arriba de los resultados la app muestra cuántos comprobantes cargó y de dónde.

**Cómo está cargado el sheet** (`comprobantes_helper.py`):
- `nombre_apellido` es el **titular de la cuenta que pagó** (puede ser una empresa o un familiar). El cliente sale de `titular_declarado` → nombre del archivo (`Comprobante_[MAIL_]Nombre_Apellido_fecha.ext`) → `nombre_apellido`, en ese orden.
- `monto` y `fecha` son texto libre del OCR en muchos formatos (`1.234.000,00`, `123.456`, `17/septiembre/2026 a las 9:17.`, `30062026`...). Si la fecha no se lee, se usa la de envío con más margen.
- Duplicados (mismo comprobante por WhatsApp y mail) se descartan.

**Cuándo se usa**:
1. CUIT que no está en deudores ni en semanas anteriores → comprobantes de ese **CUIT de pagador**.
2. Sin CUIT (`DNET CREDITO NE...` / `587 - DATANET`: transferencias desde home banking de empresa, el extracto no trae el CUIT; el número `NE…`/`CTE …` es de la operación y cambia cada vez) o CUIT sin comprobante → comprobante con **el mismo monto (±$1) y fecha** entre 4 días antes y 1 después de la acreditación.

**Cómo elige al cliente** (solo imputa si queda **uno**; si no, a ambiguos con los candidatos):
1. Alguna variante del nombre da un único cliente en deudores.
2. Si no: otros comprobantes del **mismo pagador** (mismo mail, CUIT o titular) que ya quedaron imputados en semanas anteriores (por monto + fecha) → el nombre de col H. Así, imputado a mano el primero, los siguientes de esa cuenta salen solos.
3. Si no: filas de deudores que comparten una palabra del nombre y cuyo teórico × N (N ≤ 6) es exactamente el monto. Si cierra con varios clientes, va a ambiguos con `Nombre (N × teórico)`.

Lo identificado así aparece en la columna "Identificado por" para revisarlo.

## Cómo correr la imputación de una semana

Desde la app (`streamlit run app.py` o `arrancar.bat`; en producción, la de Streamlit Cloud): se suben imputaciones y deudores, se revisa el reporte (especialmente los ambiguos) y se confirma la escritura. `imputar_core.procesar` **auto-detecta** la hoja (última "S NNN" o "USD N") y las columnas del mes (busca "MES AÑO teorico" en los headers de deudores). No hace falta tocar nada al cambiar de semana o de mes.

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
3. Si no hay historial en deudores, busca en las últimas hojas de imputaciones (semanas previas) la cuota imputada para ese CUIT. **Excepción**: si el cliente tiene otro lote con historial, la fila sin historial es un lote nuevo → `c1` (la cuota de las semanas previas es del otro lote)
4. Si aún no hay dato, usa `CUOTA_OVERRIDE` (dict manual en el CONFIG del script)

## Cómo resuelve el script clientes no encontrados por CUIT

Muchos clientes transfieren desde el número de cuenta (`402...`, `440...`), no desde su CUIT. Si el CUIT/cuenta de una transferencia no está en deudores:
1. Busca ese número en las **últimas 20 hojas** de imputaciones (6 si es USD) **y en las filas ya imputadas —amarillas— de la hoja actual** (las hojas USD son acumulativas: el antecedente del mismo cliente suele estar más abajo en la misma hoja) → extrae el nombre de col H
2. Del texto de col H se corta todo lo que va desde el primer marcador `cN`/`lN` en adelante (`"Fulano c8 en ambos lotes mismo monto"` → `"Fulano"`), porque las notas sueltas rompían la búsqueda
3. Busca ese nombre en deudores por palabras clave (case-insensitive)
4. Si hay 1 match → imputa; si hay **varios pero todos del mismo titular** → los trata como lotes del mismo cliente (sigue por la lógica de reparto); si hay 0 o varios con nombres distintos → reporta como ambiguo
5. Si tampoco hay antecedente, busca en los comprobantes del bot (ver "Comprobantes del bot"); también cuando la transferencia no trae CUIT (DATANET)
6. Dos transferencias de **cuentas distintas** que resuelven al mismo cliente por nombre se imputan como cuotas consecutivas en la misma fila (la identidad es el nombre, no el número de cuenta)

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
- Si ya fue imputado este mes (pesos, un lote): si el monto es el teórico del **mes siguiente** y esa columna está vacía → va al mes siguiente (cuota X+1; típico: paga a fin de mes la cuota que viene). Si no, y es la cuota del mes (o múltiplo) → se suma al mes actual (`=365200+373560`, `21 y 22`). Si no coincide con ninguno → ambiguo. Nunca si la fila ya tiene col H o si el pago del mes en deudores es esta misma transferencia (misma fecha y monto: la oficina la cargó sin pintar la fila)
- Al imputar: escribir en col H `{nombre} [l{lote}] c{cuota}`, en col I `x`, pintar fila amarilla

## BOLSA CEMENTO: teórico por tramo de fecha

El precio de la bolsa se congela del 1 al 10 de cada mes; después sube (sept-26: 1–10 $8.300, 11–15 $8.390, 16+ $8.490). La oficina cambiaba a mano la fórmula del teórico de cada fila según el día de pago. El script calcula el teórico él mismo:

`teórico = BOLSAS POR MES (col U) × precio del tramo del día de la transferencia × 2 si es bolsa de 50 kg`

- Tramos: encabezados de la **fila 1** del mes (`... 1 AL 10 DE SEPT 2026`, `... 10 al 15 ...`, `... desde el 16 ...`), precio en la **fila 2**. Se ignoran las columnas `los q firmaron hasta agosto 2025` (son el mismo precio ×2). Un `DESC <MES> 26` en fila 1 aplica como coeficiente. En la app se pueden sobreescribir (`10: 8300, 15: 8390, 31: 8490`).
- **Precio por día** (oct-26 en adelante): encabezados `... monto 6 y 9 oct 26`, `... monto 1/10/2026`, `... monto 7, 8, 13 y 14 OCT` en **cualquier fila arriba** de los encabezados de la tabla (fila 1, 3...), con el precio en la fila de abajo. El día de la transferencia toma el precio de ese día; un día sin precio, el del último día cargado antes. Si un día aparece en dos columnas gana la de más a la derecha. Un encabezado con tramo del mismo mes (`desde el 16 DE SEPT`) se lee como tramo aunque diga `monto`; para otro mes vale su `monto` (`... desde el 16 DE SEPT 2026 - monto 15/10/2026 en adelante` = $8.490 desde el 15/10).
- **Finde/feriado** (precio por día): si entre el último día con precio y el de acreditación hubo días sin precio y el precio cambió, un pago que cubre al menos la cuota (o N cuotas) al precio de ese día anterior va como cuota entera a ese precio (acordado con el usuario: la transferencia se hizo el sábado/feriado y se acreditó el martes).
- El tramo congelado se estira hasta el **día 13** (`DIA_TOPE_CONGELADO`): una transferencia del 10 que cae en finde/feriado aparece acreditada el 12 o 13. El mensaje de reclamo igual dice "después del día 10".
- Bolsa de 50 kg (×2): si la fórmula del teórico de la fila apunta a una columna "firmaron hasta agosto 2025"; si no se puede, fecha firma (col K) ≤ 31/08/2025.
- Otros meses / otras hojas: teórico de la planilla. Si el archivo no trae valores cacheados (lo guardó openpyxl), `LectorDeudores` evalúa las fórmulas aritméticas simples (`=U7*EX$2`, `=S7/T7`, `=497000+6000`).

## Partes de cuota (solo pesos)

Formato igual al que usa la oficina a mano. Col H: `Nombre parte de c33`, `Nombre completa c32`, `Nombre completa c5 y c6`, `Nombre c14 y completa c13`. Deudores: el "pago real" se va sumando como fórmula (`=330000+9600`), N° cuota `parte de c33` → al completarse `33` (o `32 y parte de c33`), la fecha queda la del **primer** pago (sirve para calcular el saldo al precio de ese día).

Una parte pendiente = celda "N° cuota" con `parte de cN` en los últimos 3 meses (`MESES_PARTE`) con saldo > tolerancia. También cuenta una cuota anotada **entera** (`16`) pero pagada de menos (típico: pagó el precio congelado después del 10 y la oficina la imputó entera), cuando la transferencia es justo lo que le falta (±$500, `SALDO_EXACTO_TOL`) o es un pago chico (< 90% del teórico del mes) y a esa cuota le falta poco (más que la tolerancia y hasta el 10% de su teórico; diferencias mayores se asumen aceptadas por la oficina). Si la completa → `completa c16`; si no alcanza, la celda pasa a `parte de c16` y sale el reclamo por el resto. Saldo = (cuotas que menciona la celda × teórico de ese mes) − pagado. Si se completa en un mes posterior, se suma en la columna del mes **de la parte**.

Criterio (acordado con el usuario): *si el monto parece la cuota del mes, es la cuota del mes; si no, se cancela primero lo más viejo.*

Qué hace con cada transferencia (un lote destino):
1. **Cuota entera** (±tolerancia) o múltiplo → cuota(s) siguiente(s) `X+1`, **sin mirar partes pendientes** (X = mayor cuota mencionada, incluidas las parciales).
2. **Paga de menos**:
   - sin partes pendientes, o el pago es ≥ 90% del teórico (`UMBRAL_CUOTA_DEL_MES`) y supera lo adeudado → `parte de cX+1` (la cuota del mes pagada de menos, típico: precio congelado pagado después del 10).
   - si no → completa partes pendientes (la más vieja primero); si no alcanza, acumula en esa parte; lo que sobra queda `parte de cX+1`.
3. **Paga de más**: si el sobrante alcanza para completar partes pendientes → `cX+1 y completa cN`; si no, como siempre (normal si ≤ $50.000, si no "PAGO MAS").

**Saldo a favor** (pesos, un lote): la oficina carga lo que sobró en una columna `sobrante <MES> <AÑO>` (`=EX-EW`, o `=EX-2*EW` si fueron 2 cuotas) y al cliente se le avisa que lo descuente. En el primer pago del mes, si el pago + el sobrante del **último mes con pago** ≈ N cuotas (±tolerancia), va como cuota(s) entera(s): col H `Nombre c14 (usa saldo a favor $10.000)`, pago real del mes `=pago+saldo` y el del mes del sobrante `=...-saldo` (el sobrante queda en 0, no se vuelve a usar). Se ignora si la planilla no lo respalda (pagado − cuotas × teórico de ese mes < sobrante). Si paga entero sin descontar, el saldo no se arrastra: queda en la columna del mes viejo.

**Reclamo por WhatsApp** (solo BOLSA CEMENTO): cada cuota que queda incompleta **al final de la corrida** genera un mensaje (si en la misma semana mandó la diferencia, no hay reclamo; si hubo varios pagos parciales, vale el último) con el saldo (y las otras partes abiertas) y un link `wa.me` al teléfono de la col TELEFONO (normalizado a `549…`; `telefonos_extra` permite pasarle teléfonos de otra fuente, p. ej. Supabase).

## Para cambiar de mes

No hace falta hacer nada — el script detecta el mes a partir de las fechas de las transferencias y busca la columna correspondiente en deudores automáticamente. Si el archivo de deudores no tiene ese mes aún, el script lanza un error explicativo.
