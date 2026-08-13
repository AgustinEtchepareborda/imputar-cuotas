import io
import os
import re
import traceback
import openpyxl
import streamlit as st
import pandas as pd

from imputar_core import procesar, aplicar
from comprobantes_helper import cargar_indice
from mep_helper import cargar_mep


@st.cache_data(ttl=3600)
def cargar_mep_cacheado():
    return cargar_mep()

st.set_page_config(page_title="Imputar Cuotas", page_icon="📊", layout="wide")

CACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'datos', 'comprobantes_cache.json')

# ── Helpers ──────────────────────────────────────────────────────────────────

def leer_hojas(file_bytes):
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
        names = wb.sheetnames
        wb.close()
        return names
    except Exception:
        return []

def hojas_pesos(sheets):
    return [s for s in sheets if re.match(r'^S\s*\d+', s, re.IGNORECASE)]

def hojas_usd(sheets):
    return [s for s in sheets if re.match(r'^USD\s*\d+', s, re.IGNORECASE)]

def parse_cuota_override(texto):
    resultado = {}
    if not texto.strip():
        return resultado
    for linea in texto.strip().splitlines():
        linea = linea.strip()
        if not linea or linea.startswith('#'):
            continue
        partes = re.split(r'[:\s,]+', linea)
        if len(partes) >= 2:
            cuit = re.sub(r'[^0-9]', '', partes[0])
            try:
                cuota = int(partes[1])
                if cuit:
                    resultado[cuit] = cuota
            except ValueError:
                pass
    return resultado

def fmt_monto(val, es_usd):
    if val is None:
        return ''
    if es_usd:
        return f'U$D {val:,.2f}'
    return f'${val:,.0f}'

def fmt_dif(val, es_usd):
    if val is None:
        return ''
    signo = '+' if val >= 0 else ''
    if es_usd:
        return f'{signo}{val:,.2f}'
    return f'{signo}{val:,.0f}'

# ── UI ───────────────────────────────────────────────────────────────────────

st.title('📊 Imputar Cuotas')

try:
    col1, col2 = st.columns(2)
    with col1:
        imp_file = st.file_uploader('Archivo de imputaciones (.xlsx)', type='xlsx', key='imp_upload')
    with col2:
        deu_file = st.file_uploader('Archivo de deudores (.xlsx)', type='xlsx', key='deu_upload')

    if not imp_file or not deu_file:
        st.info('Subí ambos archivos para continuar.')
        st.stop()

    imp_bytes = imp_file.read()
    deu_bytes = deu_file.read()

    hojas = leer_hojas(imp_bytes)
    if not hojas:
        st.error('No se pudo leer el archivo de imputaciones.')
        st.stop()

    h_pesos = hojas_pesos(hojas)
    h_usd = hojas_usd(hojas)

    if not h_pesos and not h_usd:
        st.error('No se encontraron hojas de semanas (formato "S 121" o "USD 5") en el archivo.')
        st.stop()

    st.divider()
    col_tipo, col_semana, col_tol, col_maxrow = st.columns([1, 2, 1, 1])

    with col_tipo:
        tipo = st.radio('Tipo', ['Pesos', 'USD'], horizontal=True)

    es_usd = tipo == 'USD'
    opciones_semana = h_usd if es_usd else h_pesos

    with col_semana:
        if not opciones_semana:
            st.warning(f'No hay hojas de tipo {"USD" if es_usd else "Pesos"} en el archivo.')
            st.stop()
        semana = st.selectbox('Semana', opciones_semana, index=len(opciones_semana) - 1)

    with col_tol:
        tolerancia = st.number_input(
            'Tolerancia',
            min_value=0,
            value=5 if es_usd else 3000,
            help='Diferencia máxima aceptable para imputar sin marcar "PAGO MENOS"'
        )

    with col_maxrow:
        max_row = st.number_input('Filas máx.', min_value=1, value=500,
                                  help='Hasta qué fila procesar en la hoja de imputaciones')

    with st.expander('Overrides manuales de cuota'):
        st.caption('Una línea por CUIT: `20111111112: 12`')
        override_text = st.text_area('CUOTA_OVERRIDE', value='', height=100, label_visibility='collapsed')

    cuota_override = parse_cuota_override(override_text)

    st.divider()

    # ── Simular ───────────────────────────────────────────────────────────────

    if st.button('🔍 Simular', type='primary', use_container_width=True):
        comprobantes_cache = cargar_indice(CACHE_PATH) if os.path.exists(CACHE_PATH) else {}
        mep_rates, mep_origen = ({}, None)
        if not es_usd:
            mep_rates, mep_origen = cargar_mep_cacheado()
        logs = []

        with st.spinner('Procesando...'):
            results, pago_menos, pago_mas, ambiguous, sin_fila, usd_en_pesos, mes_info, sheets_cfg = procesar(
                imp_bytes=imp_bytes,
                deu_bytes=deu_bytes,
                imp_sheet=semana,
                es_usd=es_usd,
                tolerance=int(tolerancia),
                max_row=int(max_row),
                cuota_override=cuota_override,
                comprobantes_cache=comprobantes_cache,
                mep_rates=mep_rates,
                log_fn=logs.append,
            )
            st.session_state['sim'] = {
                'results': results,
                'pago_menos': pago_menos,
                'pago_mas': pago_mas,
                'ambiguous': ambiguous,
                'sin_fila': sin_fila,
                'usd_en_pesos': usd_en_pesos,
                'mep_origen': mep_origen,
                'mes_info': mes_info,
                'sheets_cfg': sheets_cfg,
                'semana': semana,
                'es_usd': es_usd,
                'imp_bytes': imp_bytes,
                'deu_bytes': deu_bytes,
                'imp_name': imp_file.name,
                'deu_name': deu_file.name,
            }
            st.session_state.pop('descarga', None)

        if logs:
            with st.expander('Log de resoluciones'):
                st.code('\n'.join(logs))

    # ── Mostrar resultados ────────────────────────────────────────────────────

    sim = st.session_state.get('sim')
    if not sim:
        st.stop()

    results   = sim['results']
    pago_menos = sim['pago_menos']
    pago_mas   = sim.get('pago_mas', [])
    ambiguous  = sim['ambiguous']
    es_usd_sim = sim['es_usd']

    mes_vals = list(sim['mes_info'].values())
    mes_label = mes_vals[0] if mes_vals else '?'
    st.caption(f'Mes detectado en deudores: **{mes_label}**')

    usd_en_pesos = sim.get('usd_en_pesos', [])

    total = len(results) + len(pago_menos) + len(pago_mas) + len(ambiguous) + len(usd_en_pesos)
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric('Total filas', total)
    c2.metric('✅ Para imputar', len(results))
    c3.metric('💵 USD en pesos', len(usd_en_pesos))
    c4.metric('⚠️ Pago menos', len(pago_menos))
    c5.metric('🔺 Pago más', len(pago_mas))
    c6.metric('❌ Ambiguos', len(ambiguous))

    if results:
        st.subheader(f'✅ Para imputar ({len(results)})')
        rows_ok = []
        for r in results:
            rows_ok.append({
                'Fila': r['imp_row'],
                'Cliente': f"{r['cliente']}{r['lote_str']}",
                'CUIT': r['cuit'],
                'Monto real': fmt_monto(r['monto_real'], es_usd_sim),
                'Teórico': fmt_monto(r['monto_teo'], es_usd_sim),
                'Diferencia': fmt_dif(r['diferencia'], es_usd_sim),
                'Cuota': r['cuota'],
                'Fecha': r['fecha'].strftime('%d/%m/%Y') if r['fecha'] else '?',
                'Hoja deudores': r['hoja'],
            })
        st.dataframe(pd.DataFrame(rows_ok), use_container_width=True, hide_index=True)

    usd_aprobados_filas = []
    if usd_en_pesos:
        st.subheader(f'💵 Clientes USD fijo que pagaron en pesos ({len(usd_en_pesos)})')
        origen = sim.get('mep_origen')
        st.caption(
            f'Cotización dólar MEP venta del día de la transferencia (fuente: misma serie que Ámbito Financiero'
            f'{", " + origen if origen else ""}). Tildá **Imputar** en los que quieras escribir: '
            'en USD fijo va el monto en pesos como "pago real", más N° de cuota y fecha.'
        )
        rows_usd = []
        for e in usd_en_pesos:
            rows_usd.append({
                'Imputar': False,
                'Fila': e['imp_row'],
                'Cliente': str(e['cliente']),
                'CUIT': e['cuit'],
                'Fecha': e['fecha'].strftime('%d/%m/%Y') if e['fecha'] else '?',
                'Pagó en $': fmt_monto(e['monto_pesos'], False),
                'MEP venta': f"${e['mep']:,.2f}" if e['mep'] else 'sin dato',
                'Equiv. U$D': fmt_monto(e['equiv_usd'], True) if e['equiv_usd'] is not None else '?',
                'Teórico U$D': fmt_monto(e['teo_usd'], True) if e['teo_usd'] is not None else '?',
                'Dif. U$D': fmt_dif(e['dif_usd'], True) if e['dif_usd'] is not None else '?',
                'Cuota': e['cuota'],
            })
        df_usd = pd.DataFrame(rows_usd)
        edited_usd = st.data_editor(
            df_usd,
            use_container_width=True,
            hide_index=True,
            disabled=[c for c in df_usd.columns if c != 'Imputar'],
            key=f"usd_editor_{sim['semana']}",
        )
        usd_aprobados_filas = [int(f) for f, ok in zip(edited_usd['Fila'], edited_usd['Imputar']) if ok]

    if pago_menos:
        st.subheader(f'⚠️ Pago menos ({len(pago_menos)}) — se escribirá "PAGO MENOS" en col H')
        rows_pm = []
        for p in pago_menos:
            rows_pm.append({
                'Fila': p['row'],
                'Cliente': p['cliente'],
                'CUIT': p['cuit'],
                'Transferido': fmt_monto(p['transferido'], es_usd_sim),
                'Teórico': fmt_monto(p['teorico'], es_usd_sim),
                'Diferencia': fmt_monto(p['diferencia'], es_usd_sim),
            })
        st.dataframe(pd.DataFrame(rows_pm), use_container_width=True, hide_index=True)

    if pago_mas:
        st.subheader(f'🔺 Pago más ({len(pago_mas)}) — se escribirá "PAGO MAS" en col H')
        rows_pmas = []
        for p in pago_mas:
            rows_pmas.append({
                'Fila': p['row'],
                'Cliente': p['cliente'],
                'CUIT': p['cuit'],
                'Transferido': fmt_monto(p['transferido'], es_usd_sim),
                'Teórico': fmt_monto(p['teorico'], es_usd_sim),
                'Diferencia': fmt_dif(p['diferencia'], es_usd_sim),
            })
        st.dataframe(pd.DataFrame(rows_pmas), use_container_width=True, hide_index=True)

    if ambiguous:
        st.subheader(f'❌ Casos ambiguos ({len(ambiguous)}) — revisar manualmente')
        rows_amb = []
        for a in ambiguous:
            cliente = a.get('cliente', a.get('concepto', ''))
            monto = a.get('monto', '')
            monto_str = fmt_monto(monto, es_usd_sim) if isinstance(monto, (int, float)) else str(monto) if monto else ''
            matches_str = str(a.get('matches', '')) if a.get('matches') else ''
            rows_amb.append({
                'Fila': a.get('row', '?'),
                'Motivo': a.get('motivo', ''),
                'Cliente / Concepto': str(cliente)[:60] if cliente else '',
                'Monto': monto_str,
                'Candidatos': matches_str[:80] if matches_str else '',
            })
        st.dataframe(pd.DataFrame(rows_amb), use_container_width=True, hide_index=True)

    # ── Confirmar e imputar ───────────────────────────────────────────────────

    if not results and not pago_menos and not usd_en_pesos:
        st.info('No hay nada para imputar.')
        st.stop()

    st.divider()
    aviso = f'Esto escribirá **{len(results)}** imputaciones y **{len(pago_menos)}** "PAGO MENOS" en los archivos.'
    if usd_en_pesos:
        aviso += f' USD en pesos habilitados: **{len(usd_aprobados_filas)}** de {len(usd_en_pesos)}.'
    st.warning(aviso)

    if st.button('✅ Confirmar e Imputar', type='primary', use_container_width=True):
        usd_aprobados = [e for e in usd_en_pesos if e['imp_row'] in usd_aprobados_filas]
        with st.spinner('Escribiendo archivos...'):
            imp_out, deu_out = aplicar(
                results=sim['results'],
                pago_menos=sim['pago_menos'],
                pago_mas=sim.get('pago_mas', []),
                sin_fila=sim.get('sin_fila', []),
                usd_en_pesos=usd_aprobados,
                imp_bytes=sim['imp_bytes'],
                deu_bytes=sim['deu_bytes'],
                imp_sheet=sim['semana'],
                sheets_cfg=sim['sheets_cfg'],
            )
            st.session_state['descarga'] = {
                'imp_out': imp_out,
                'deu_out': deu_out,
                'imp_name': sim['imp_name'],
                'deu_name': sim['deu_name'],
            }

    descarga = st.session_state.get('descarga')
    if descarga:
        st.success('Archivos listos. Descargalos y reemplazá los originales en la carpeta compartida.')
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button(
                '⬇️ Descargar imputaciones.xlsx',
                data=descarga['imp_out'],
                file_name=descarga['imp_name'],
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True,
            )
        with col_d2:
            st.download_button(
                '⬇️ Descargar deudores.xlsx',
                data=descarga['deu_out'],
                file_name=descarga['deu_name'],
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True,
            )

except Exception as e:
    st.error('Error en la aplicación:')
    st.code(traceback.format_exc())
