# 📊 Imputar Cuotas

Herramienta interna para **automatizar la imputación de cuotas** a partir de las transferencias bancarias semanales. Toma el archivo de imputaciones (transferencias del banco) y registra en la planilla de deudores el monto real, el número de cuota y la fecha de cada pago, identificando al cliente por su CUIT.

Cuenta con una interfaz web ([Streamlit](https://streamlit.io/)) para operar sin tocar código, y scripts de línea de comandos para uso avanzado.

> ⚠️ **Repositorio privado.** Procesa datos personales de clientes (CUITs, nombres, montos). No debe hacerse público.

---

## ✨ Características

- **Identificación automática de clientes** por CUIT extraído del concepto de la transferencia (soporta guiones, múltiples CUITs por celda y texto libre).
- **Detección automática** de la hoja semanal (`S 121`, `USD 5`, etc.) y de las columnas del mes en la planilla de deudores: no hay que reconfigurar al cambiar de semana o de mes.
- **Cálculo del número de cuota** a partir del historial del cliente, con posibilidad de override manual.
- **Manejo de casos borde**: pagos de menos (tolerancia configurable), múltiples lotes por cliente, clientes no encontrados por CUIT (búsqueda por nombre).
- **Modo _dry-run_**: permite revisar el reporte completo antes de escribir en la planilla.
- Soporte para cuotas en **pesos** y en **dólares** (con cotización MEP).

---

## 🛠️ Tecnologías

- Python 3
- [Streamlit](https://streamlit.io/) — interfaz web
- [openpyxl](https://openpyxl.readthedocs.io/) — lectura/escritura de Excel
- [pandas](https://pandas.pydata.org/) — manejo de datos
- [gspread](https://docs.gspread.org/) — acceso a Google Sheets (cache de comprobantes)

---

## 📂 Estructura del proyecto

```
├── app.py                    # Interfaz web (Streamlit)
├── imputar_core.py           # Lógica central de imputación
├── imputar_s120.py           # Script CLI — imputación en pesos
├── imputar_usd5.py           # Script CLI — imputación en dólares
├── comprobantes_helper.py    # Acceso al cache de comprobantes (Google Sheets)
├── exportar_comprobantes.py  # Refresca el cache desde Google Sheets
├── mep_helper.py             # Cotización del dólar MEP
├── requirements.txt          # Dependencias
├── arrancar.bat              # Lanzador rápido (Windows)
└── .streamlit/config.toml    # Configuración de Streamlit
```

---

## 🚀 Instalación

Requiere **Python 3.11+**.

```bash
# 1. (Opcional) crear entorno virtual
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # Linux/Mac

# 2. Instalar dependencias
pip install -r requirements.txt
```

---

## ▶️ Uso

### Interfaz web (recomendado)

```bash
streamlit run app.py
```

En Windows también podés hacer doble clic en **`arrancar.bat`**.

Desde la web: se suben los archivos de imputaciones y de deudores, se revisa el reporte y se confirma la escritura.

### Línea de comandos

Los scripts corren por defecto en modo `DRY_RUN = True` (solo reportan, no escriben):

```bash
python imputar_s120.py      # pesos
python imputar_usd5.py      # dólares
```

Flujo recomendado:

1. Correr con `DRY_RUN = True` y revisar el reporte (especialmente los casos ambiguos).
2. Cambiar a `DRY_RUN = False` y volver a correr para escribir en la planilla.

---

## 🔒 Configuración y datos sensibles

- El **ID del Google Sheet** de comprobantes se toma de la variable de entorno `COMPROBANTES_SHEET_ID` (no está en el código):

  ```bash
  set COMPROBANTES_SHEET_ID=tu_id_de_sheet      # Windows
  # export COMPROBANTES_SHEET_ID=tu_id_de_sheet # Linux/Mac
  ```

- Las credenciales y secretos van en `.streamlit/secrets.toml` (ignorado por Git).
- Los archivos de datos (`.xlsx`, cache de comprobantes) **no se versionan**.
- El acceso a Google Sheets usa OAuth: la primera ejecución abre el navegador para autorizar.

---

## 📌 Notas

Proyecto de uso interno. Para el detalle de las reglas de negocio (cálculo de cuota, tolerancias, manejo de lotes, estructura de las planillas) ver [`CLAUDE.md`](CLAUDE.md).
