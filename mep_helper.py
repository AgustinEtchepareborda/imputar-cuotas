"""
Cotización histórica del dólar MEP (bolsa), la misma que publica Ámbito Financiero.

Fuente: https://api.argentinadatos.com/v1/cotizaciones/dolares/bolsa
(API pública y gratuita con el historial completo día por día, compra y venta).

Se guarda un cache local en datos/mep_cache.json para poder trabajar sin
conexión con los últimos valores descargados.
"""

import json
import os
import datetime
import urllib.request

API_URL = 'https://api.argentinadatos.com/v1/cotizaciones/dolares/bolsa'
CACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'datos', 'mep_cache.json')


def _descargar():
    req = urllib.request.Request(API_URL, headers={'User-Agent': 'Mozilla/5.0'})
    data = json.load(urllib.request.urlopen(req, timeout=30))
    return {d['fecha']: d['venta'] for d in data if d.get('venta')}


def cargar_mep(cache_path=CACHE_PATH):
    """Retorna (rates, origen) donde rates es dict 'YYYY-MM-DD' -> venta.

    Intenta descargar el historial completo; si no hay conexión usa el último
    cache guardado. origen es 'online', 'cache' o 'sin datos'.
    """
    try:
        rates = _descargar()
        try:
            with open(cache_path, 'w', encoding='utf-8') as f:
                json.dump(rates, f)
        except OSError:
            pass
        return rates, 'online'
    except Exception:
        try:
            with open(cache_path, encoding='utf-8') as f:
                return json.load(f), 'cache'
        except (OSError, ValueError):
            return {}, 'sin datos'


def mep_para_fecha(rates, fecha, max_dias_atras=4):
    """Valor venta del MEP para una fecha (datetime, date o 'YYYY-MM-DD').

    Si la fecha exacta no tiene cotización (feriado), usa el último cierre
    anterior hasta max_dias_atras días. Retorna (valor, fecha_usada) o (None, None).
    """
    if isinstance(fecha, datetime.datetime):
        fecha = fecha.date()
    elif isinstance(fecha, str):
        try:
            fecha = datetime.date.fromisoformat(fecha)
        except ValueError:
            return None, None
    if not isinstance(fecha, datetime.date):
        return None, None
    for delta in range(max_dias_atras + 1):
        clave = (fecha - datetime.timedelta(days=delta)).isoformat()
        if clave in rates:
            return rates[clave], clave
    return None, None


if __name__ == '__main__':
    rates, origen = cargar_mep()
    print(f'{len(rates)} cotizaciones cargadas ({origen})')
    if rates:
        ultima = max(rates)
        print(f'Última: {ultima} -> venta {rates[ultima]}')
