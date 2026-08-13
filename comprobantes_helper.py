"""
Helper para el cache local del Google Sheet "Datos comprobantes".

Flujo:
  1. Antes de imputar, correr: python exportar_comprobantes.py
     → descarga el sheet y actualiza datos/comprobantes_cache.json
  2. Los scripts imputar_*.py cargan el cache automáticamente como fallback.
"""

import json
import re
import os

# ID del Google Sheet de comprobantes. Se toma de la variable de entorno
# COMPROBANTES_SHEET_ID (o de .streamlit/secrets.toml) para no versionarlo.
SPREADSHEET_ID = os.environ.get('COMPROBANTES_SHEET_ID', '')
DEFAULT_CACHE   = 'datos/comprobantes_cache.json'


def _normalize_cuit(raw):
    if not raw:
        return None
    digits = re.sub(r'[^0-9]', '', str(raw))
    return digits if len(digits) >= 10 else None


def _nombre_desde_filename(filename):
    """'Comprobante_Apellido_Nombre_2026-05-01_1234.pdf' → 'Apellido Nombre'"""
    if not filename:
        return None
    name = re.sub(r'^Comprobante_', '', str(filename), flags=re.IGNORECASE)
    name = re.sub(r'_\d{4}-\d{2}-\d{2}_\d+\.(pdf|jpg|jpeg|png)$', '', name, flags=re.IGNORECASE)
    return name.replace('_', ' ').strip() or None


def cargar_indice(cache_path=DEFAULT_CACHE):
    """Lee el cache local. Devuelve dict {cuit_digits: nombre}."""
    if not os.path.exists(cache_path):
        return {}
    with open(cache_path, encoding='utf-8') as f:
        return json.load(f)


def exportar_cache(cache_path=DEFAULT_CACHE):
    """
    Descarga el sheet de comprobantes vía gspread y actualiza el cache local.
    Requiere: pip install gspread
    Primera vez abre el navegador para autorizar; luego queda guardado.
    """
    try:
        import gspread
    except ImportError:
        raise SystemExit('Falta instalar gspread: pip install gspread')

    if not SPREADSHEET_ID:
        raise SystemExit(
            'Falta configurar el ID del sheet. Definir la variable de entorno '
            'COMPROBANTES_SHEET_ID antes de correr.')

    print('Conectando con Google Sheets...')
    gc = gspread.oauth()
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.sheet1
    rows = ws.get_all_values()
    print(f'  -> {len(rows) - 1} filas descargadas')

    indice = {}
    for row in rows[1:]:  # omitir encabezado
        nombre_col = row[1].strip() if len(row) > 1 else ''
        cuit_raw   = row[2].strip() if len(row) > 2 else ''
        filename   = row[8].strip() if len(row) > 8 else ''

        cuit = _normalize_cuit(cuit_raw)
        if not cuit:
            continue

        nombre = nombre_col or _nombre_desde_filename(filename)
        if nombre and cuit not in indice:
            indice[cuit] = nombre

    os.makedirs(os.path.dirname(os.path.abspath(cache_path)), exist_ok=True)
    with open(cache_path, 'w', encoding='utf-8') as f:
        json.dump(indice, f, ensure_ascii=False, indent=2)
    print(f'  -> Cache actualizado: {len(indice)} CUITs en {cache_path}')
    return indice
