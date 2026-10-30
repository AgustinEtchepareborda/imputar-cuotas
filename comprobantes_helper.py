"""
Acceso al Google Sheet "Datos comprobantes" (lo llena el bot de WhatsApp y el
de mail con los comprobantes que mandan los clientes).

Se usa como último recurso para identificar transferencias que no se pudieron
asignar por CUIT (ni en deudores ni en semanas anteriores):
  * por el CUIT de quien pagó (col cuit_cuil), o
  * si la transferencia no trae CUIT (DATANET, Interbanking) o el CUIT no
    aparece, por monto exacto + fecha cercana.
El sheet solo dice QUIÉN es el cliente; el nombre después se busca en deudores.

Cómo está cargado el sheet (ojo):
  * nombre_apellido es el TITULAR DE LA CUENTA QUE PAGÓ (puede ser una empresa
    o un familiar), no necesariamente el cliente. El cliente sale primero de
    titular_declarado y si no del nombre del archivo
    (Comprobante_[MAIL_]Nombre_Apellido_YYYY-MM-DD_HHMM[_N].ext).
  * monto y fecha son texto libre leído del comprobante, en muchos formatos.
  * El mismo comprobante puede estar dos veces (WhatsApp + mail).

Origen de los datos (cargar_filas):
  1. Google Sheets con cuenta de servicio (secrets: gcp_service_account +
     COMPROBANTES_SHEET_ID).
  2. Cache local datos/comprobantes_cache.json (python exportar_comprobantes.py).
"""

import datetime
import json
import os
import re

# ID del Google Sheet de comprobantes. Se toma de la variable de entorno
# COMPROBANTES_SHEET_ID (o de .streamlit/secrets.toml) para no versionarlo.
SPREADSHEET_ID = os.environ.get('COMPROBANTES_SHEET_ID', '')
DEFAULT_CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'datos', 'comprobantes_cache.json')
SECRETS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.streamlit', 'secrets.toml')

# Ventana de fechas para matchear por monto: el comprobante puede ser de unos
# días antes de la acreditación (finde, feriado, DATANET acredita al otro día).
DIAS_ANTES = 4
DIAS_DESPUES = 1
# Si la fecha del comprobante no se pudo leer se usa la de envío: más margen.
DIAS_DESPUES_ENVIO = 3

_MESES = {
    'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6, 'jul': 7,
    'ago': 8, 'sep': 9, 'set': 9, 'oct': 10, 'nov': 11, 'dic': 12,
}


def _normalize_cuit(raw):
    if not raw:
        return None
    s = str(raw).strip()
    if re.fullmatch(r'\d+\.0', s):  # 20331032647.0 (celda numérica)
        s = s[:-2]
    digits = re.sub(r'[^0-9]', '', s)
    return digits if len(digits) >= 10 else None


def _nombre_desde_filename(filename):
    """'Comprobante_MAIL_Apellido_Nombre_2026-05-01_1234_0.pdf' → 'Apellido Nombre'"""
    if not filename:
        return None
    name = re.sub(r'\.(pdf|jpe?g|png|webp|heic)$', '', str(filename).strip(), flags=re.IGNORECASE)
    name = re.sub(r'^Comprobante_', '', name, flags=re.IGNORECASE)
    name = re.sub(r'^MAIL_', '', name, flags=re.IGNORECASE)
    name = re.sub(r'_\d{4}-\d{2}-\d{2}(_\d+)*$', '', name)
    name = name.replace('_', ' ').strip()
    # nombres que son solo un teléfono o un número no sirven
    return name if re.search(r'[A-Za-zÁÉÍÓÚáéíóúÑñ]{3,}', name) else None


def parse_monto(raw):
    """Monto del comprobante en formato argentino (o lo que el OCR haya dejado).
    '1.234.000,00' → 1234000.0, '123.456' → 123456.0, '-111.200,00' → 111200.0,
    '100,00' → 100.0, '123,456' → 123456.0. None si es ambiguo."""
    if raw is None:
        return None
    s = str(raw).strip()
    s = re.sub(r'\s(\d{2})$', r',\1', s)  # '123.456 00' → '123.456,00'
    s = re.sub(r'[^0-9.,]', '', s)
    if not s or not re.search(r'\d', s):
        return None
    if '.' in s and ',' in s:
        dec = max(s.rfind('.'), s.rfind(','))
        entero = re.sub(r'[.,]', '', s[:dec])
        s = f'{entero}.{s[dec + 1:]}'
    elif ',' in s:
        partes = s.split(',')
        if len(partes) > 2:
            return None
        s = partes[0] + ('' if len(partes[1]) == 3 else '.') + partes[1]
    elif '.' in s:
        partes = s.split('.')
        if all(len(p) == 3 for p in partes[1:]):
            s = ''.join(partes)
        elif len(partes) > 2 and len(partes[-1]) == 2 and all(len(p) == 3 for p in partes[1:-1]):
            s = ''.join(partes[:-1]) + '.' + partes[-1]  # '123.000.00'
        elif len(partes) == 2 and len(partes[1]) <= 2:
            pass
        else:
            return None
    try:
        return abs(float(s))
    except ValueError:
        return None


def parse_fecha(raw):
    """Fecha del comprobante (texto libre) → datetime.date o None.
    Formatos vistos: '10/09/2026', '2026-09-10 13:15', '17/septiembre/2026 a
    las 9:17.', 'Lunes, 03 de agosto de 2026 a las 10:12 hs', '04/SEP/2026',
    '30062026', '10/09/26'."""
    if raw is None:
        return None
    s = str(raw).strip().lower()
    if not s:
        return None
    m = re.search(r'(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m = re.search(r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{4}|\d{2})\b', s)
        if m:
            d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        else:
            m = re.search(r'(\d{1,2})(?:\s+de\s+|\s*[/.-]\s*|\s+)([a-záéíóú]{3,})\.?,?(?:\s+del?\s+|\s*[/.-]\s*|\s+)(\d{4})', s)
            if m and m.group(2)[:3] in _MESES:
                d, mo, y = int(m.group(1)), _MESES[m.group(2)[:3]], int(m.group(3))
            else:
                m = re.fullmatch(r'(\d{2})(\d{2})(\d{4})(\.0)?', s)
                if not m:
                    return None
                d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y < 100:
        y += 2000
    try:
        return datetime.date(y, mo, d)
    except ValueError:
        return None


def _clave_nombre(nombre):
    """'EMPRESA  Ejemplo S.A.' → 'empresa ejemplo s.a.' (None si es muy corto
    para identificar a alguien)."""
    s = re.sub(r'\s+', ' ', str(nombre or '')).strip().lower()
    return s if len(s.split()) >= 2 else None


class Comprobantes:
    """Índice de los comprobantes del sheet, listo para buscar por CUIT o por
    monto + fecha. Cada registro:
      {'nombres': [...variantes del nombre del cliente, en orden de confianza],
       'pagador', 'cuit', 'monto', 'moneda', 'fecha' (date), 'fecha_exacta',
       'mz_lote', 'archivo', 'id'}"""

    def __init__(self, registros=None):
        self.registros = registros or []
        self.por_cuit = {}
        for r in self.registros:
            if r.get('cuit'):
                self.por_cuit.setdefault(r['cuit'], []).append(r)

    def __len__(self):
        return len(self.registros)

    @property
    def n_cuits(self):
        return len(self.por_cuit)

    @classmethod
    def desde_filas(cls, filas):
        """filas: lista de listas con la fila de encabezados primero (lo que
        devuelve get_all_values() del sheet)."""
        if not filas:
            return cls()
        header = [str(h).strip().lower() for h in filas[0]]

        def col(row, nombre):
            i = header.index(nombre) if nombre in header else None
            return str(row[i]).strip() if i is not None and i < len(row) and row[i] is not None else ''

        registros, vistos = [], set()
        for row in filas[1:]:
            pagador = col(row, 'nombre_apellido')
            nombres = []
            for n in (col(row, 'titular_declarado'), _nombre_desde_filename(col(row, 'filename')), pagador):
                if n and n.lower() not in [x.lower() for x in nombres]:
                    nombres.append(n)
            if not nombres:
                continue
            fecha = parse_fecha(col(row, 'fecha'))
            fecha_exacta = fecha is not None
            if fecha is None:
                fecha = parse_fecha(col(row, 'timestamp'))
            monto = parse_monto(col(row, 'monto'))
            moneda = col(row, 'moneda').upper()
            moneda = moneda if moneda in ('ARS', 'USD') else ''
            # el mismo comprobante mandado por WhatsApp y por mail
            id_c = col(row, 'id_comprobante')
            clave = (id_c if id_c and not id_c.upper().startswith('HASH') else None,
                     monto, fecha, nombres[0].lower())
            if clave in vistos:
                continue
            vistos.add(clave)
            registros.append({
                'nombres': nombres,
                'pagador': pagador,
                'cuit': _normalize_cuit(col(row, 'cuit_cuil')),
                'monto': monto,
                'moneda': moneda,
                'fecha': fecha,
                'fecha_exacta': fecha_exacta,
                'mz_lote': col(row, 'manzana_lote'),
                'archivo': col(row, 'filename'),
                'id': id_c,
                'remitente': col(row, 'remitente').lower(),
            })
        return cls(registros)

    @classmethod
    def desde_dict_legacy(cls, d):
        """Cache viejo {cuit: nombre}: solo sirve para buscar por CUIT."""
        return cls([{'nombres': [n], 'pagador': n, 'cuit': c, 'monto': None, 'moneda': '',
                     'fecha': None, 'fecha_exacta': False, 'mz_lote': '', 'archivo': '', 'id': '',
                     'remitente': ''}
                    for c, n in (d or {}).items() if n])

    def por_cuit_de(self, cuit):
        """Registros de ese CUIT de pagador, el más reciente primero."""
        regs = self.por_cuit.get(cuit, [])
        return sorted(regs, key=lambda r: r['fecha'] or datetime.date.min, reverse=True)

    def del_mismo_pagador(self, reg):
        """Otros comprobantes (con monto y fecha) mandados desde el mismo mail,
        con el mismo CUIT de pagador o a nombre del mismo titular."""
        pagador = _clave_nombre(reg.get('pagador'))
        out = []
        for r in self.registros:
            if r is reg or r['monto'] is None or r['fecha'] is None:
                continue
            if ((reg.get('remitente') and r.get('remitente') == reg['remitente'])
                    or (reg.get('cuit') and r.get('cuit') == reg['cuit'])
                    or (pagador and _clave_nombre(r.get('pagador')) == pagador)):
                out.append(r)
        return out

    def por_monto(self, monto, fecha, moneda, tol=1.0):
        """Registros con el mismo monto (±tol) y fecha compatible con la de la
        transferencia (`fecha`: date/datetime)."""
        if not monto or fecha is None:
            return []
        if isinstance(fecha, datetime.datetime):
            fecha = fecha.date()
        out = []
        for r in self.registros:
            if r['monto'] is None or r['fecha'] is None:
                continue
            if moneda and r['moneda'] and r['moneda'] != moneda:
                continue
            if abs(r['monto'] - monto) > tol:
                continue
            despues = DIAS_DESPUES if r['fecha_exacta'] else DIAS_DESPUES_ENVIO
            dias = (fecha - r['fecha']).days
            if -despues <= dias <= DIAS_ANTES:
                out.append(r)
        return out


def describir(reg):
    """Texto corto para el log/reporte: 'Pagador (archivo.jpg) $1.234.000 01/01'."""
    monto = f" ${reg['monto']:,.0f}".replace(',', '.') if reg.get('monto') else ''
    fecha = f" {reg['fecha']:%d/%m}" if reg.get('fecha') else ''
    archivo = f" ({reg['archivo']})" if reg.get('archivo') else ''
    return f"{reg.get('pagador') or reg['nombres'][0]}{archivo}{monto}{fecha}"


# ── Carga de datos ────────────────────────────────────────────────────────────

def leer_sheet(service_account_info, sheet_id):
    """Lee la hoja 1 del sheet con una cuenta de servicio. Devuelve las filas
    (lista de listas, encabezado incluido)."""
    import gspread
    gc = gspread.service_account_from_dict(dict(service_account_info))
    return gc.open_by_key(sheet_id).sheet1.get_all_values()


def _secrets_locales():
    if not os.path.exists(SECRETS_PATH):
        return {}
    import tomllib
    with open(SECRETS_PATH, 'rb') as f:
        return tomllib.load(f)


def cargar_filas(secrets=None, cache_path=DEFAULT_CACHE):
    """(filas, origen). Prueba el sheet con la cuenta de servicio y si no, el
    cache local. `secrets`: mapping con gcp_service_account y
    COMPROBANTES_SHEET_ID (st.secrets en la app); si es None se lee
    .streamlit/secrets.toml."""
    secrets = _secrets_locales() if secrets is None else secrets
    sa = secrets.get('gcp_service_account') if secrets else None
    sheet_id = (secrets.get('COMPROBANTES_SHEET_ID') if secrets else None) or SPREADSHEET_ID
    error = None
    if sa and sheet_id:
        try:
            return leer_sheet(sa, sheet_id), 'Google Sheets'
        except Exception as e:  # sin red, permiso revocado, etc. → cache
            error = f'{type(e).__name__}: {e}'
    if os.path.exists(cache_path):
        with open(cache_path, encoding='utf-8') as f:
            data = json.load(f)
        if isinstance(data, dict):  # formato viejo {cuit: nombre}
            return data, 'cache local (formato viejo)'
        return data, 'cache local' + (f' — no se pudo leer el sheet ({error})' if error else '')
    return [], (f'no se pudo leer el sheet ({error})' if error else 'sin datos')


def cargar_comprobantes(secrets=None, cache_path=DEFAULT_CACHE):
    """(Comprobantes, origen)."""
    filas, origen = cargar_filas(secrets, cache_path)
    if isinstance(filas, dict):
        return Comprobantes.desde_dict_legacy(filas), origen
    return Comprobantes.desde_filas(filas), origen


def cargar_indice(cache_path=DEFAULT_CACHE):
    """Índice simple {cuit: nombre} (lo usan los drivers locales)."""
    comp, _ = cargar_comprobantes(cache_path=cache_path)
    return {c: comp.por_cuit_de(c)[0]['nombres'][0] for c in comp.por_cuit}


def exportar_cache(cache_path=DEFAULT_CACHE):
    """Descarga el sheet y lo guarda en el cache local (para correr sin red).
    Usa la cuenta de servicio de .streamlit/secrets.toml; si no hay, OAuth."""
    secrets = _secrets_locales()
    sheet_id = secrets.get('COMPROBANTES_SHEET_ID') or SPREADSHEET_ID
    if not sheet_id:
        raise SystemExit(
            'Falta configurar el ID del sheet: COMPROBANTES_SHEET_ID en '
            '.streamlit/secrets.toml o como variable de entorno.')
    print('Conectando con Google Sheets...')
    if secrets.get('gcp_service_account'):
        filas = leer_sheet(secrets['gcp_service_account'], sheet_id)
    else:
        try:
            import gspread
        except ImportError:
            raise SystemExit('Falta instalar gspread: pip install gspread')
        filas = gspread.oauth().open_by_key(sheet_id).sheet1.get_all_values()
    print(f'  -> {len(filas) - 1} filas descargadas')
    os.makedirs(os.path.dirname(os.path.abspath(cache_path)), exist_ok=True)
    with open(cache_path, 'w', encoding='utf-8') as f:
        json.dump(filas, f, ensure_ascii=False)
    comp = Comprobantes.desde_filas(filas)
    print(f'  -> Cache actualizado: {len(comp)} comprobantes, {comp.n_cuits} CUITs en {cache_path}')
    return comp
