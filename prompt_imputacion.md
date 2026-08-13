# Prompt de imputación de cuotas

Este es el prompt original usado para automatizar la imputación semanal.
Guardado para mejoras futuras.

---

Quiero que actúes como asistente de imputación de cuotas trabajando sobre los archivos Excel de este proyecto.

Archivos:

* `imputaciones.xlsx`
* `deudores.xlsx`

Tu tarea es analizar las filas pendientes en la hoja de S 120 de `imputaciones.xlsx` y completar automáticamente la información correspondiente en `deudores.xlsx`. Fijate en la S 119 para ver como esta imputado todo, en deudores no quiero que toques todavia la pagina de usd fijo.

# Reglas

## 1. Filas pendientes

Procesar únicamente filas que:

* NO estén marcadas en amarillo
* estén pendientes de imputación

Las filas amarillas ya fueron procesadas y no deben tocarse.

---

## 2. Identificación de cliente

Buscar coincidencias usando:

* CUIT/CUIL

Si la fila de imputaciones no tiene CUIT/CUIL:

* no imputar
* dejar la fila igual

---

## 3. Datos a completar en deudores.xlsx

Para cada imputación válida completar:

* monto real
* fecha
* número de cuota

La fecha y el monto real salen directamente de `imputaciones.xlsx`.

---

## 4. Número de cuota

Determinar el próximo número de cuota buscando:

* la última cuota previamente imputada del cliente
* y sumando 1

---

## 5. Diferencia de monto

Comparar:

* monto transferido
* vs monto teórico

Si el cliente pagó menos y la diferencia es MAYOR a 3000 pesos:

* NO imputar
* escribir:

```txt
PAGO MENOS
```

en la fila correspondiente de `imputaciones.xlsx`

* dejar la fila sin marcar en amarillo

Si la diferencia es menor o igual a 3000:

* imputar normalmente

---

## 6. Coincidencias múltiples

Si una transferencia coincide con múltiples clientes:

* imputar al cliente cuyo monto teórico coincida con el monto transferido

Si más de un cliente tiene el mismo monto teórico:

* no imputar
* dejar la fila pendiente

---

## 7. Actualización de imputaciones.xlsx

Después de imputar correctamente:

* escribir nombre del cliente
* escribir número de cuota
* marcar fila en amarillo

---

## 8. Seguridad

Antes de modificar archivos:

* analizar la estructura de ambos Excel
* identificar correctamente columnas y formato
* avisar cualquier ambigüedad o problema

Si no estás seguro de una imputación:

* NO imputar automáticamente
* reportar el caso

---

## Objetivo

Quiero minimizar trabajo manual pero evitar imputaciones incorrectas.

Priorizá precisión antes que automatizar todo.
