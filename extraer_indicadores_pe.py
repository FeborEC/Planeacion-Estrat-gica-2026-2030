# -*- coding: utf-8 -*-
"""
Extracción de indicadores del Plan Estratégico 2026-2030 desde las fuentes de información.

Lee las fuentes mensuales que ya usa el Modelo de Riesgos Financieros (balance de la hoja «Balance» de
Solvencia_PT_Mes, repositorios de cartera e individual de aportes), calcula las variables de captura de
las hojas IND-XX de Tablero_Indicadores_PE.xlsx y:

  1. Siempre genera Comparacion_Extraccion_PE.xlsx: por hoja, variable y mes, lo digitado frente a lo
     calculado, la diferencia y qué se hizo con la celda.
  2. Según el MODO, escribe en las celdas azules (de captura) del tablero:
       COMPARAR   no escribe nada (solo el informe)
       VACIOS     llena solo las celdas vacías (no toca lo ya digitado)
       REEMPLAZAR sobrescribe todas las celdas calculables con lo de las fuentes
     Las celdas con fórmula nunca se tocan.

La escritura se hace con el propio Excel (una instancia aparte, invisible): así se conservan comentarios,
formatos y fórmulas, y Excel recalcula y guarda los resultados que lee Plan_Estrategico_Dashboard.py.
Antes de escribir se saca un respaldo del tablero en la carpeta _respaldos.

Todo lo configurable (rutas, mes inicial, modo, cuentas contables y reglas de cada variable) está en
Parametros_Extraccion_PE.xlsx; si no existe, el script lo crea con los valores iniciales.

Uso:
    python extraer_indicadores_pe.py                 usa el MODO del libro de parámetros
    python extraer_indicadores_pe.py --comparar      solo el informe de comparación
    python extraer_indicadores_pe.py --vacios        llena las celdas vacías
    python extraer_indicadores_pe.py --reemplazar    sobrescribe las celdas calculables
    python extraer_indicadores_pe.py --reemplazar --hojas IND-13    solo escribe en las hojas indicadas
"""
import datetime
import glob
import os
import re
import shutil
import sys
import unicodedata

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

VERSION = '1.1.1'
BASE = os.path.dirname(os.path.abspath(__file__))
TABLERO = os.path.join(BASE, 'Tablero_Indicadores_PE.xlsx')
PARAMS = os.path.join(BASE, 'Parametros_Extraccion_PE.xlsx')
INFORME = os.path.join(BASE, 'Comparacion_Extraccion_PE.xlsx')
RESPALDOS = os.path.join(BASE, '_respaldos')

MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre',
         'Noviembre', 'Diciembre']
ABREV = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
CAL = ['A', 'B', 'C', 'D', 'E']


class ErrorExtraccion(Exception):
    """Error con un mensaje pensado para quien ejecuta el script."""


# ============================================================================================================
# 1. PARÁMETROS (Parametros_Extraccion_PE.xlsx)
# ============================================================================================================
GENERAL = [
    ['RUTA_FUENTES', r'D:\Diana Fonseca\Asistente Claude\Modelo de Riesgos Global\Fuentes de información',
     'Carpeta «Fuentes de información» del Modelo de Riesgos (Solvencia_PT_Mes, Repositorios_Cartera_Mes, Individual_Aportes_Mes)'],
    ['RUTA_PARAMETROS_RIESGOS', r'D:\Diana Fonseca\Asistente Claude\Modelo de Riesgos Global\Fuentes de información\Parametros\Parametros_Riesgos_Financieros.xlsx',
     'Libro de parámetros del Modelo de Riesgos: de su hoja PRODUCTOS sale la categoría de cada línea de crédito (Rotativo, Tarjeta, Consumo…)'],
    ['MES_INICIAL', '2026-01', 'Primer mes (AAAA-MM) que se calcula; también se lee diciembre del año anterior como base'],
    ['MODO', 'VACIOS', 'COMPARAR = solo informe · VACIOS = llena solo celdas vacías · REEMPLAZAR = sobrescribe las celdas calculables (las fórmulas nunca se tocan)'],
    ['MESES_PROMEDIO', 12, 'FUENTE BALANCE_PROMEDIO: número de cierres mensuales que se promedian (el mes y los anteriores; 12 = el mes y los 11 anteriores)'],
    ['TOLERANCIA_RELATIVA', 0.0001, 'En el informe, diferencias relativas hasta este valor se marcan como «Coincide» (0,0001 = 0,01 %)'],
]

# HOJA, COLUMNA (texto del encabezado en la fila de captura), FUENTE, DEFINICION, ACTIVO, DESCRIPCION
# FUENTE:
#   BALANCE_MES    movimiento del mes de las cuentas (estado de resultados: saldo acumulado del mes menos el del mes anterior; en enero, el acumulado)
#   BALANCE_SALDO  saldo de las cuentas al cierre del mes (en resultados = acumulado del año)
#   BALANCE_PROMEDIO promedio del saldo de las cuentas en los últimos MESES_PROMEDIO cierres (el mes y los anteriores)
#   CARTERA        TOTAL · CATEGORIAS:cat1;cat2 · MORA_OBJETIVA (calificación objetiva B a E) · LINEA:nombre · LINEA_A:nombre (calificación objetiva A)
#                  · TRANSICION_RECUPERACION · TRANSICION_DETERIORO (matriz de transición objetiva mes anterior → mes, como el Modelo de Riesgos)
#   ASOCIADOS      TOTAL · BASE_DIC · SE_MANTIENEN · NUEVOS · RETIROS (cédulas del individual de aportes frente a diciembre del año anterior)
# DEFINICION de balance: cuentas con signo y factor opcional, p. ej. «51+52», «21-210595-211095», «0.2*3505»
VARIABLES = [
    ['IND-F1', 'Gastos operativos totales ($)', 'BALANCE_MES', '51+52', 'SI', 'Gastos de administración (51) y otros gastos (52) del mes'],
    ['IND-F1', 'Ingresos totales ($)', 'BALANCE_MES', '4', 'SI', 'Ingresos (4) del mes'],
    ['IND-F2', 'Costos financieros totales ($)', 'BALANCE_MES', '6', 'SI', 'Costo de ventas (6): intereses de depósitos y créditos, del mes'],
    ['IND-F2', 'Ingresos totales ($)', 'BALANCE_MES', '4', 'SI', 'Ingresos (4) del mes'],
    ['IND-F3', 'Excedente del período ($)', 'BALANCE_MES', '53', 'SI', 'Excedentes y pérdidas del ejercicio (53) del mes'],
    ['IND-F3', 'Ingresos totales ($)', 'BALANCE_MES', '4', 'SI', 'Ingresos (4) del mes'],
    ['IND-01', 'Saldo cartera líneas mayor rendimiento ($)', 'CARTERA', 'CATEGORIAS:Rotativo;Tarjeta', 'SI', 'Saldo de capital de las líneas de categoría Rotativo y Tarjeta (hoja PRODUCTOS del Modelo de Riesgos)'],
    ['IND-01', 'Saldo total de cartera ($)', 'CARTERA', 'TOTAL', 'SI', 'Saldo de capital de todo el repositorio de cartera'],
    ['IND-02', 'Ingresos de cartera del mes ($)', 'BALANCE_MES', '4150', 'SI', 'Ingresos cartera de créditos (4150) del mes'],
    ['IND-02', 'Ingresos Acumulados', 'BALANCE_SALDO', '4150', 'SI', 'Ingresos cartera de créditos (4150) acumulados del año'],
    ['IND-03', 'Ingresos por comisiones del mes ($)', 'BALANCE_MES', '424090005', 'SI', 'Comisiones compañías de seguros (424090005) del mes'],
    ['IND-03', 'Ingresos Acumulados', 'BALANCE_SALDO', '424090005', 'SI', 'Comisiones compañías de seguros (424090005) acumuladas del año'],
    ['IND-05', 'Saldo de cartera en mora ≥ B ($)', 'CARTERA', 'MORA_OBJETIVA', 'SI', 'Saldo de capital en calificación objetiva B a E'],
    ['IND-05', 'Saldo total de cartera ($)', 'CARTERA', 'TOTAL', 'SI', 'Saldo de capital de todo el repositorio de cartera'],
    ['IND-06', 'Saldo cartera recuperada ($)', 'CARTERA', 'TRANSICION_RECUPERACION', 'SI', 'Matriz objetiva: saldo que pasó a una calificación mejor más lo cancelado que venía en B a E'],
    ['IND-06', 'Saldo cartera deteriorada ($)', 'CARTERA', 'TRANSICION_DETERIORO', 'SI', 'Matriz objetiva: saldo actual de los créditos que pasaron a una calificación peor'],
    ['IND-07', 'Saldo normalizaciones al día ($)', 'CARTERA', 'LINEA_A:117 - NORMALIZACION DE CARTERA', 'SI', 'Saldo de la línea de normalización en calificación objetiva A'],
    ['IND-07', 'Saldo total normalizaciones ($)', 'CARTERA', 'LINEA:117 - NORMALIZACION DE CARTERA', 'SI', 'Saldo total de la línea de normalización'],
    ['IND-08', 'Saldo total de depósitos del mes ($)', 'BALANCE_SALDO', '21-210595-211095-212595-213095', 'SI', 'Depósitos (21) sin intereses por pagar (210595, 211095, 212595, 213095); también llena la base de diciembre del año anterior'],
    ['IND-10', 'Margen financiero bruto ($)', 'BALANCE_SALDO', '4150-6150', 'SI', 'Ingresos de cartera (4150) menos intereses de depósitos (6150), acumulados del año'],
    ['IND-10', 'Gastos operativos ($)', 'BALANCE_SALDO', '51-4225', 'SI', 'Gastos de administración (51) netos de recuperaciones de deterioro (4225), acumulados del año'],
    ['IND-12', 'Aportes amortizados ($)', 'BALANCE_SALDO', '311010', 'SI', 'Aportes amortizados (311010)'],
    ['IND-12', 'Reserva protección aportes ($)', 'BALANCE_SALDO', '3205', 'SI', 'Reserva protección de aportes (3205)'],
    ['IND-12', 'Excedente del mes ($)', 'BALANCE_SALDO', '0.2*3505', 'SI', '20 % del excedente acumulado del año (3505), parte que va a la reserva'],
    ['IND-12', 'Resultado de ejercicios anteriore', 'BALANCE_SALDO', '0.2*3905', 'SI', '20 % del excedente de ejercicios anteriores pendiente de distribuir (3905)'],
    ['IND-12', 'Activo total ($)', 'BALANCE_SALDO', '1', 'SI', 'Activo (1)'],
    ['IND-13', 'Excedente ($)', 'BALANCE_SALDO', '53', 'SI', 'Excedente acumulado del año (53)'],
    ['IND-13', 'Patrimonio Promedio ($)', 'BALANCE_PROMEDIO', '3', 'SI', 'Promedio del patrimonio (3) de los últimos MESES_PROMEDIO cierres mensuales (el mes y los 11 anteriores)'],
    ['IND-14', 'Total de asociados al cierre del mes', 'ASOCIADOS', 'TOTAL', 'SI', 'Cédulas del individual de aportes del mes'],
    ['IND-15', 'Base social (dic año anterior )', 'ASOCIADOS', 'BASE_DIC', 'SI', 'Cédulas del individual de aportes de diciembre del año anterior'],
    ['IND-15', 'Base social actual', 'ASOCIADOS', 'TOTAL', 'SI', 'Cédulas del individual de aportes del mes'],
    ['IND-15', 'Base social que se mantiene (estaban antes y ahora)', 'ASOCIADOS', 'SE_MANTIENEN', 'SI', 'Cédulas que están en diciembre del año anterior y en el mes'],
    ['IND-15', 'Asociados nuevos', 'ASOCIADOS', 'NUEVOS', 'SI', 'Cédulas del mes que no estaban en diciembre del año anterior'],
    ['IND-15', 'Retiros Totales', 'ASOCIADOS', 'RETIROS', 'SI', 'Cédulas de diciembre del año anterior que ya no están en el mes'],
]


def crear_parametros():
    wb = openpyxl.Workbook()
    cab = Font(bold=True, color='FFFFFF')
    fondo = PatternFill('solid', fgColor='0B5D3B')
    ws = wb.active
    ws.title = 'GENERAL'
    ws.append(['PARAMETRO', 'VALOR', 'DESCRIPCION'])
    for r in GENERAL:
        ws.append(r)
    ws2 = wb.create_sheet('VARIABLES')
    ws2.append(['HOJA', 'COLUMNA', 'FUENTE', 'DEFINICION', 'ACTIVO', 'DESCRIPCION'])
    for r in VARIABLES:
        ws2.append(r)
    ws3 = wb.create_sheet('AYUDA')
    for t in [
        'Parámetros de extraer_indicadores_pe.py (Plan Estratégico 2026-2030)',
        '',
        'GENERAL: rutas de las fuentes, primer mes, modo de escritura y tolerancia del informe.',
        'VARIABLES: una fila por columna de captura de una hoja IND-XX del tablero. COLUMNA = texto del encabezado de la fila de captura (fila «#  Año  Mes …»).',
        'FUENTE = BALANCE_MES: movimiento del mes (saldo acumulado del mes menos el del mes anterior; en enero, el acumulado).',
        'FUENTE = BALANCE_SALDO: saldo al cierre del mes (en cuentas de resultados es el acumulado del año).',
        'FUENTE = BALANCE_PROMEDIO: promedio del saldo de las cuentas en los últimos MESES_PROMEDIO cierres mensuales (hoja GENERAL; 12 = el mes y los 11 anteriores).',
        '   DEFINICION: cuentas con signo y factor opcional, p. ej. 51+52, 21-210595-211095, 0.2*3505. Toma el NUEVO SALDO de la hoja «Balance» de Solvencia_PT_Mes.',
        'FUENTE = CARTERA (repositorio de cartera, saldo de capital): TOTAL, CATEGORIAS:Rotativo;Tarjeta, MORA_OBJETIVA, LINEA:nombre de la línea, LINEA_A:nombre de la línea (solo calificación objetiva A),',
        '   TRANSICION_RECUPERACION y TRANSICION_DETERIORO (matriz de transición objetiva del mes anterior al mes, igual que el Modelo de Riesgos).',
        'FUENTE = ASOCIADOS (cédulas del individual de aportes): TOTAL, BASE_DIC, SE_MANTIENEN, NUEVOS, RETIROS (frente a diciembre del año anterior).',
        'ACTIVO = SI / NO. Una fila en NO no se calcula ni se escribe.',
        'Las celdas con fórmula del tablero nunca se tocan. Antes de escribir se saca un respaldo en la carpeta _respaldos.',
    ]:
        ws3.append([t])
    for w in (ws, ws2):
        for c in w[1]:
            c.font, c.fill = cab, fondo
        w.freeze_panes = 'A2'
    for col, an in zip('ABC', (28, 100, 110)):
        ws.column_dimensions[col].width = an
    for col, an in zip('ABCDEF', (10, 46, 16, 40, 8, 110)):
        ws2.column_dimensions[col].width = an
    ws3.column_dimensions['A'].width = 180
    wb.save(PARAMS)


def leer_parametros():
    if not os.path.exists(PARAMS):
        crear_parametros()
        print(f'Se creó {os.path.basename(PARAMS)} con los valores iniciales.')
    g = pd.read_excel(PARAMS, sheet_name='GENERAL')
    gen = {str(a).strip(): b for a, b in zip(g['PARAMETRO'], g['VALOR']) if pd.notna(a)}
    v = pd.read_excel(PARAMS, sheet_name='VARIABLES')
    v = v[v['ACTIVO'].astype(str).str.strip().str.upper().isin(['SI', 'SÍ'])]
    return gen, [r for r in v.to_dict('records') if pd.notna(r['HOJA']) and pd.notna(r['COLUMNA'])]


# ============================================================================================================
# 2. FUENTES
# ============================================================================================================
def _am_de(y, n):
    return y * 100 + n


def _archivo(raiz, carpeta, prefijo, am):
    y, m = divmod(am, 100)
    c = glob.glob(os.path.join(raiz, carpeta, str(y), f'{m:03d}_{prefijo}_*.xls*'))
    c = [x for x in c if not os.path.basename(x).startswith('~$')]
    return c[0] if c else None


def meses_a_calcular(raiz, mes_inicial):
    """Desde diciembre del año anterior al mes inicial hasta el último mes con balance."""
    y0, m0 = (int(x) for x in str(mes_inicial)[:7].split('-'))
    am = _am_de(y0 - 1, 12)
    out = []
    while True:
        if not _archivo(raiz, 'Solvencia_PT_Mes', 'SolvPT', am):
            if am > _am_de(y0, m0):
                break
        out.append(am)
        y, m = divmod(am, 100)
        am = _am_de(y + (m == 12), 1 if m == 12 else m + 1)
        if am > 209912:
            break
    return out


def leer_balance(raiz, am):
    f = _archivo(raiz, 'Solvencia_PT_Mes', 'SolvPT', am)
    if not f:
        return None
    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
    if 'Balance' not in wb.sheetnames:
        raise ErrorExtraccion(f'{os.path.basename(f)} no tiene la hoja «Balance»')
    d = {}
    for r in wb['Balance'].iter_rows(min_row=3, values_only=True):
        try:
            d[int(float(r[3]))] = float(r[8] or 0)          # D = CUENTA, I = NUEVO SALDO
        except (TypeError, ValueError, IndexError):
            continue
    wb.close()
    return d


def productos_categoria(ruta):
    try:
        p = pd.read_excel(ruta, sheet_name='PRODUCTOS')
        cat = {}
        for a, b in zip(p['NOMBRE_PRODUCTO'], p['CATEGORIA']):
            if pd.notna(a) and pd.notna(b):
                cat.setdefault(_n(a), str(b).strip())          # como BUSCARV y el Modelo de Riesgos: vale la primera coincidencia
        return cat
    except Exception as e:                                   # noqa: BLE001
        print(f'   ! No se pudo leer la hoja PRODUCTOS ({e}); las categorías de línea quedan vacías')
        return {}


def leer_cartera(raiz, am, cat):
    f = _archivo(raiz, 'Repositorios_Cartera_Mes', 'Rep', am)
    if not f:
        return None
    d = pd.read_excel(f, usecols=['NUMERO_PRESTAMO', 'NOMBRE_PRODUCTO', 'CALIFICACION_OBJETIVA', 'SALDO_CAPITAL'])
    d['SALDO_CAPITAL'] = pd.to_numeric(d['SALDO_CAPITAL'], errors='coerce').fillna(0.0)
    d['CRED'] = d['NUMERO_PRESTAMO'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    d['LINEA'] = d['NOMBRE_PRODUCTO'].astype(str).str.strip()
    d['CAT'] = d['LINEA'].map(_n).map(cat).fillna('')
    d['CAL'] = d['CALIFICACION_OBJETIVA'].astype(str).str.strip().str.upper().str[:1].map({c: i for i, c in enumerate(CAL)})
    return d


def leer_asociados(raiz, am):
    f = _archivo(raiz, 'Individual_Aportes_Mes', 'Ap', am)
    if not f:
        return None
    d = pd.read_excel(f, usecols=['IDENTIFICATION'])
    return set(d['IDENTIFICATION'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True))


# ============================================================================================================
# 3. CÁLCULOS
# ============================================================================================================
TOKEN = re.compile(r'([+-]?)\s*(?:(\d+(?:[.,]\d+)?)\s*\*\s*)?(\d+)')


def eval_cuentas(defi, bal):
    defi = str(defi).replace(' ', '')
    if not defi or TOKEN.sub('', defi):
        raise ErrorExtraccion(f'Definición de cuentas no válida: «{defi}» (use p. ej. 51+52, 21-210595, 0.2*3505)')
    t = 0.0
    for sg, fac, cta in TOKEN.findall(defi):
        t += (-1 if sg == '-' else 1) * (float(fac.replace(',', '.')) if fac else 1.0) * bal.get(int(cta), 0.0)
    return t


def transicion_objetiva(prev, act):
    """Recuperación y deterioro de la matriz objetiva mes anterior → mes (misma regla del Modelo de Riesgos):
    créditos con saldo y calificación A-E; deterioro = saldo actual de los que pasan a peor calificación;
    recuperación = saldo actual de los que pasan a mejor calificación + saldo anterior de los cancelados que venían en B-E."""
    p = prev[(prev['SALDO_CAPITAL'] > 0) & prev['CAL'].notna()].groupby('CRED').agg(a=('CAL', 'first'), s=('SALDO_CAPITAL', 'sum'))
    a = act[(act['SALDO_CAPITAL'] > 0) & act['CAL'].notna()].groupby('CRED').agg(x=('CAL', 'first'), s=('SALDO_CAPITAL', 'sum'))
    j = p.join(a, how='outer', lsuffix='_p', rsuffix='_a')
    sigue = j['a'].notna() & j['x'].notna()
    de = j.loc[sigue & (j['x'] > j['a']), 's_a'].sum()
    re_ = j.loc[sigue & (j['x'] < j['a']), 's_a'].sum() + j.loc[j['a'].notna() & j['x'].isna() & (j['a'] >= 1), 's_p'].sum()
    return float(re_), float(de)


def ventana(am, n):
    """El mes y los n-1 cierres anteriores, del más antiguo al más reciente."""
    out, (y, m) = [am], divmod(am, 100)
    for _ in range(n - 1):
        y, m = (y - 1, 12) if m == 1 else (y, m - 1)
        out.append(_am_de(y, m))
    return out[::-1]


def calcular(gen, variables):
    raiz = str(gen.get('RUTA_FUENTES', ''))
    if not os.path.isdir(raiz):
        raise ErrorExtraccion(f'No existe la carpeta de fuentes: {raiz} (parámetro RUTA_FUENTES)')
    meses = meses_a_calcular(raiz, gen.get('MES_INICIAL', '2026-01'))
    fuentes = {str(v['FUENTE']).strip().upper() for v in variables}
    print(f'Meses: {meses[0]} a {meses[-1]}')
    bal, car, aso = {}, {}, {}
    cat = productos_categoria(str(gen.get('RUTA_PARAMETROS_RIESGOS', ''))) if 'CARTERA' in fuentes else {}
    nprom = int(float(gen.get('MESES_PROMEDIO', 12) or 12))
    if 'BALANCE_PROMEDIO' in fuentes:                     # cierres anteriores al primer mes que pide el promedio
        y, m = divmod(meses[0], 100)
        for _ in range(nprom - 1):
            y, m = (y - 1, 12) if m == 1 else (y, m - 1)
            bal[_am_de(y, m)] = leer_balance(raiz, _am_de(y, m))
    for am in meses:
        if fuentes & {'BALANCE_MES', 'BALANCE_SALDO', 'BALANCE_PROMEDIO'}:
            bal[am] = leer_balance(raiz, am)
        if 'CARTERA' in fuentes:
            car[am] = leer_cartera(raiz, am, cat)
        if 'ASOCIADOS' in fuentes:
            aso[am] = leer_asociados(raiz, am)
        print(f'   {am}: balance {"sí" if bal.get(am) else "no"} · cartera {"sí" if car.get(am) is not None else "no"} · aportes {"sí" if aso.get(am) is not None else "no"}')
    sin_cat = sorted({x for d in car.values() if d is not None for x in d.loc[d['CAT'] == '', 'LINEA'].unique()})
    if sin_cat:
        print(f'   ! Líneas sin categoría en la hoja PRODUCTOS (no cuentan en CATEGORIAS): {", ".join(sin_cat)}')

    def ant(am):
        y, m = divmod(am, 100)
        return _am_de(y - 1, 12) if m == 1 else am - 1

    tr = {}
    res = {}                                   # (hoja, columna) -> {am: valor}
    for v in variables:
        hoja, col, fu, defi = str(v['HOJA']).strip(), str(v['COLUMNA']), str(v['FUENTE']).strip().upper(), str(v['DEFINICION']).strip()
        out = {}
        for am in meses:
            y, m = divmod(am, 100)
            val = None
            if fu == 'BALANCE_PROMEDIO':
                ven = ventana(am, nprom)
                val = sum(eval_cuentas(defi, bal[a]) for a in ven) / nprom if all(bal.get(a) for a in ven) else None
            elif fu in ('BALANCE_MES', 'BALANCE_SALDO') and bal.get(am):
                val = eval_cuentas(defi, bal[am])
                if fu == 'BALANCE_MES' and m != 1:
                    if not bal.get(ant(am)):
                        val = None
                    else:
                        val -= eval_cuentas(defi, bal[ant(am)])
            elif fu == 'CARTERA' and car.get(am) is not None:
                d = car[am]
                k, _, arg = defi.partition(':')
                k = k.strip().upper()
                if k == 'TOTAL':
                    val = d['SALDO_CAPITAL'].sum()
                elif k == 'CATEGORIAS':
                    cs = {x.strip().lower() for x in arg.split(';') if x.strip()}
                    val = d.loc[d['CAT'].str.lower().isin(cs), 'SALDO_CAPITAL'].sum()
                elif k == 'MORA_OBJETIVA':
                    val = d.loc[d['CAL'].between(1, 4), 'SALDO_CAPITAL'].sum()
                elif k == 'LINEA':
                    val = d.loc[d['LINEA'] == arg.strip(), 'SALDO_CAPITAL'].sum()
                elif k == 'LINEA_A':
                    val = d.loc[(d['LINEA'] == arg.strip()) & (d['CAL'] == 0), 'SALDO_CAPITAL'].sum()
                elif k in ('TRANSICION_RECUPERACION', 'TRANSICION_DETERIORO'):
                    if car.get(ant(am)) is not None:
                        if am not in tr:
                            tr[am] = transicion_objetiva(car[ant(am)], d)
                        val = tr[am][0 if k == 'TRANSICION_RECUPERACION' else 1]
                else:
                    raise ErrorExtraccion(f'{hoja} · {col}: definición de CARTERA desconocida «{defi}»')
            elif fu == 'ASOCIADOS' and aso.get(am) is not None:
                s, b = aso[am], aso.get(_am_de(y - 1, 12))
                k = defi.upper()
                if k == 'TOTAL':
                    val = len(s)
                elif b is not None and k in ('BASE_DIC', 'SE_MANTIENEN', 'NUEVOS', 'RETIROS'):
                    val = {'BASE_DIC': len(b), 'SE_MANTIENEN': len(s & b), 'NUEVOS': len(s - b), 'RETIROS': len(b - s)}[k]
                elif b is None and k in ('BASE_DIC', 'SE_MANTIENEN', 'NUEVOS', 'RETIROS'):
                    val = None
                else:
                    raise ErrorExtraccion(f'{hoja} · {col}: definición de ASOCIADOS desconocida «{defi}»')
            elif fu not in ('BALANCE_MES', 'BALANCE_SALDO', 'BALANCE_PROMEDIO', 'CARTERA', 'ASOCIADOS'):
                raise ErrorExtraccion(f'{hoja} · {col}: FUENTE desconocida «{fu}»')
            if val is not None:
                out[am] = round(float(val), 2)
        res[(hoja, col)] = out
    return res, {'meses': meses, 'bal': bal, 'car': car, 'aso': aso, 'raiz': raiz, 'nprom': nprom}


# ============================================================================================================
# 4. UBICACIÓN EN EL TABLERO Y ESCRITURA
# ============================================================================================================
def _n(t):
    t = unicodedata.normalize('NFKD', str(t or '')).encode('ascii', 'ignore').decode().lower()
    return re.sub(r'\s+', ' ', t).strip()


def ubicar(wb, hoja, col, base=True):
    """Fila de encabezado (la de «#» en la columna B), columna cuyo encabezado coincide y filas (año, mes)."""
    if hoja not in wb.sheetnames:
        raise ErrorExtraccion(f'El tablero no tiene la hoja {hoja}')
    ws = wb[hoja]
    hr = next((r for r in range(1, 60) if str(ws.cell(r, 2).value).strip() == '#'), None)
    if not hr:
        raise ErrorExtraccion(f'{hoja}: no se encontró la fila de captura («#» en la columna B)')
    cc = next((c for c in range(1, ws.max_column + 1) if _n(ws.cell(hr, c).value) == _n(col)), None)
    if not cc:
        raise ErrorExtraccion(f'{hoja}: no hay una columna con el encabezado «{col}» en la fila {hr}')
    filas = {}
    # filas de captura (debajo del encabezado); por encima, solo la fila base de diciembre del año anterior, solo en la
    # primera columna de captura (E), que es donde el tablero guarda esa base (p. ej. IND-08, E19), y solo para variables de saldo
    for r in range(1, ws.max_row + 1):
        if r <= hr and (cc != 5 or not base):
            continue
        y, mes = ws.cell(r, 3).value, ws.cell(r, 4).value
        if isinstance(y, (int, float)) and _n(mes) in [_n(x) for x in MESES]:
            filas[_am_de(int(y), [_n(x) for x in MESES].index(_n(mes)) + 1)] = r
    return cc, filas


def preparar(res, modo, tol, variables=()):
    fu = {(str(v['HOJA']).strip(), str(v['COLUMNA'])): str(v['FUENTE']).strip().upper() for v in variables}
    wbf = openpyxl.load_workbook(TABLERO)                    # fórmulas
    wbv = openpyxl.load_workbook(TABLERO, data_only=True)    # valores
    informe, escribir = [], []
    for (hoja, col), serie in res.items():
        cc, filas = ubicar(wbf, hoja, col, base=fu.get((hoja, col)) == 'BALANCE_SALDO')
        for am, calc in sorted(serie.items()):
            r = filas.get(am)
            if not r:
                continue
            f = wbf[hoja].cell(r, cc).value
            dig = wbv[hoja].cell(r, cc).value
            es_f = isinstance(f, str) and f.startswith('=')
            dign = dig if isinstance(dig, (int, float)) and not isinstance(dig, bool) else None
            dif = None if dign is None else calc - dign
            rel = None if dign is None else (abs(dif) / abs(dign) if dign else (0 if abs(dif) < 0.005 else 1))
            if es_f:
                acc = 'Fórmula: no se toca'
            elif dign is None and dig not in (None, ''):
                acc = 'Tiene texto: no se toca' if modo != 'REEMPLAZAR' else 'Se escribe (reemplaza texto)'
            elif dign is None:
                acc = 'Se escribe (estaba vacía)' if modo in ('VACIOS', 'REEMPLAZAR') else 'Vacía (modo comparar)'
            elif rel <= tol:
                acc = 'Coincide'
            else:
                acc = 'Se reemplaza' if modo == 'REEMPLAZAR' else 'Difiere: se deja lo digitado'
            if acc.startswith('Se escribe') or acc == 'Se reemplaza':
                escribir.append((hoja, r, cc, calc))
            informe.append([hoja, col, am // 100, MESES[am % 100 - 1], dign if dign is not None else dig, calc, dif, rel, acc,
                            wbf[hoja].cell(r, cc).coordinate])
    return informe, escribir


def escribir_excel(escribir):
    lock = os.path.join(BASE, '~$' + os.path.basename(TABLERO))
    if os.path.exists(lock):
        raise ErrorExtraccion('El tablero está abierto en Excel (existe ' + os.path.basename(lock) + '): ciérrelo y vuelva a ejecutar.')
    os.makedirs(RESPALDOS, exist_ok=True)
    resp = os.path.join(RESPALDOS, f'Tablero_Indicadores_PE_antes_extraccion_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx')
    shutil.copy2(TABLERO, resp)
    print(f'Respaldo: {resp}')
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx('Excel.Application')     # instancia propia: solo se cierra esta
    try:
        xl.Visible = False
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Open(TABLERO, UpdateLinks=0)
        for hoja, r, c, v in escribir:
            wb.Worksheets(hoja).Cells(r, c).Value = v
        xl.CalculateFull()
        wb.Save()
        wb.Close(False)
    finally:
        xl.Quit()
        del xl
        pythoncom.CoUninitialize()


def guardar_informe(informe, modo, escribir):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Comparacion'
    ws.append(['HOJA', 'VARIABLE', 'AÑO', 'MES', 'DIGITADO', 'CALCULADO', 'DIFERENCIA', 'DIF_RELATIVA', 'ACCION', 'CELDA'])
    for r in informe:
        ws.append(r)
    cab, fondo = Font(bold=True, color='FFFFFF'), PatternFill('solid', fgColor='0B5D3B')
    for c in ws[1]:
        c.font, c.fill, c.alignment = cab, fondo, Alignment(horizontal='center', vertical='center', wrap_text=True)
    colores = {'Coincide': 'D3F9D8', 'Difiere': 'FFE3E3', 'Se reemplaza': 'FFE9A8', 'Se escribe': 'DDE7FF'}
    for row in ws.iter_rows(min_row=2):
        for k, col in colores.items():
            if str(row[8].value).startswith(k):
                row[8].fill = PatternFill('solid', fgColor=col)
        for j in (4, 5, 6):
            row[j].number_format = '#,##0.00'
        row[7].number_format = '0.000%'
    for col, an in zip('ABCDEFGHIJ', (9, 46, 7, 12, 20, 20, 18, 12, 30, 8)):
        ws.column_dimensions[col].width = an
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    r = wb.create_sheet('Resumen')
    from collections import Counter
    cn = Counter(x[8] for x in informe)
    r.append(['Generado', f'{datetime.datetime.now():%Y-%m-%d %H:%M}'])
    r.append(['Versión del extractor', VERSION])
    r.append(['Modo', modo])
    r.append(['Celdas escritas', len(escribir)])
    r.append([])
    r.append(['ACCION', 'CELDAS'])
    for k, v in cn.most_common():
        r.append([k, v])
    r.column_dimensions['A'].width = 34
    wb.save(INFORME)


# ============================================================================================================
# 5. LIBRO DE AUDITORÍA (Auditoria_Extraccion_PE.xlsx): datos de origen y todas las variables con fórmulas de Excel
# ============================================================================================================
AUDITORIA = os.path.join(BASE, 'Auditoria_Extraccion_PE.xlsx')


def _col(i):
    s = ''
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _etm(am):
    return f'{ABREV[am % 100 - 1]}-{am // 100}'


def generar_auditoria(variables, datos, res, informe):
    """Libro formulado para auditoría. Hojas:
    Leeme · Fuentes (archivo de cada mes) · Balance (NUEVO SALDO de las cuentas usadas, por mes) · Var_Balance (fórmulas sobre Balance)
    · Cartera_Datos (un crédito por fila y mes) · Var_Cartera (SUMAR.SI.CONJUNTO) · Transicion_Datos (un crédito por par de meses, con el
    movimiento calculado por fórmula) · Var_Transicion · Asociados_Datos (una cédula por fila y mes, con fórmula «estaba en diciembre»)
    · Asociados_Dic (cédulas de diciembre y si siguen en cada mes) · Var_Asociados · Resumen (todas las variables por mes) · Comparacion
    (digitado en el tablero frente a la fórmula)."""
    meses, bal, car, aso, raiz = datos['meses'], datos['bal'], datos['car'], datos['aso'], datos['raiz']
    vb = [v for v in variables if str(v['FUENTE']).strip().upper() in ('BALANCE_MES', 'BALANCE_SALDO', 'BALANCE_PROMEDIO')]
    nprom = datos.get('nprom', 12)
    mb = sorted(a for a in bal if bal.get(a))                 # meses de la hoja Balance (incluye los anteriores del promedio)
    vc = [v for v in variables if str(v['FUENTE']).strip().upper() == 'CARTERA']
    va = [v for v in variables if str(v['FUENTE']).strip().upper() == 'ASOCIADOS']
    wb = openpyxl.Workbook()
    cab, fondo = Font(bold=True, color='FFFFFF'), PatternFill('solid', fgColor='0B5D3B')
    sub = PatternFill('solid', fgColor='E8F5EE')

    def encab(ws, fila=1):
        for c in ws[fila]:
            c.font, c.fill, c.alignment = cab, fondo, Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.freeze_panes = ws.cell(fila + 1, 1)

    def anchos(ws, a):
        for i, w in enumerate(a, 1):
            ws.column_dimensions[_col(i)].width = w

    # ---- Leeme
    ws = wb.active
    ws.title = 'Leeme'
    for t in [
        'Auditoría de la extracción de indicadores · Plan Estratégico 2026-2030',
        f'Generado {datetime.datetime.now():%Y-%m-%d %H:%M} con extraer_indicadores_pe.py v{VERSION} · meses {_etm(meses[0])} a {_etm(meses[-1])}',
        '',
        'Cada variable que el extractor escribe en Tablero_Indicadores_PE.xlsx está aquí calculada con fórmulas de Excel sobre los datos de origen:',
        '  Balance → hoja «Balance» de Solvencia_PT_Mes (NUEVO SALDO por cuenta). Var_Balance suma las cuentas de la definición; las de FUENTE BALANCE_MES restan el mes anterior (enero = acumulado).',
        '  Cartera_Datos → repositorio de cartera (un crédito por fila y mes; CATEGORIA según la hoja PRODUCTOS del Modelo de Riesgos). Var_Cartera usa SUMAR.SI.CONJUNTO.',
        '  Transicion_Datos → un crédito por par de meses (anterior → actual) con calificación objetiva y saldo de cada mes; el MOVIMIENTO es fórmula.',
        '     Deterioro = saldo actual de los que pasan a peor calificación; Recuperación = saldo actual de los que pasan a mejor + saldo anterior de los cancelados que venían en B a E.',
        '  Asociados_Datos / Asociados_Dic → cédulas del individual de aportes; «estaba en diciembre» y «sigue en el mes» son fórmulas COINCIDIR.',
        '  Resumen → todas las variables por mes (enlazadas a las hojas Var_*). Comparacion → lo digitado en el tablero frente a la fórmula, con diferencia.',
        '',
        'Definiciones editables en Parametros_Extraccion_PE.xlsx (hoja VARIABLES).',
    ]:
        ws.append([t])
    ws['A1'].font = Font(bold=True, size=14, color='0B5D3B')
    ws.column_dimensions['A'].width = 170

    # ---- Fuentes
    ws = wb.create_sheet('Fuentes')
    ws.append(['MES', 'BALANCE (Solvencia_PT_Mes)', 'REPOSITORIO DE CARTERA', 'INDIVIDUAL DE APORTES'])
    for am in meses:
        ws.append([_etm(am)] + [os.path.basename(_archivo(raiz, c, p_, am) or '—') for c, p_ in
                                (('Solvencia_PT_Mes', 'SolvPT'), ('Repositorios_Cartera_Mes', 'Rep'), ('Individual_Aportes_Mes', 'Ap'))])
    encab(ws)
    anchos(ws, (12, 30, 30, 30))

    # ---- Balance
    ctas = sorted({int(c) for v in vb for _, _, c in TOKEN.findall(str(v['DEFINICION']).replace(' ', ''))})
    nomb = {}
    for am in mb:
        f = _archivo(raiz, 'Solvencia_PT_Mes', 'SolvPT', am)
        if f and not nomb:
            w0 = openpyxl.load_workbook(f, read_only=True, data_only=True)
            for r in w0['Balance'].iter_rows(min_row=3, values_only=True):
                try:
                    if int(float(r[3])) in ctas:
                        nomb[int(float(r[3]))] = str(r[4]).strip()
                except (TypeError, ValueError, IndexError):
                    pass
            w0.close()
    ws = wb.create_sheet('Balance')
    ws.append(['CUENTA', 'NOMBRE'] + [_etm(am) for am in mb])
    fila_cta = {}
    for i, c in enumerate(ctas):
        ws.append([c, nomb.get(c, '')] + [bal[am].get(c, 0.0) for am in mb])
        fila_cta[c] = i + 2
    encab(ws)
    anchos(ws, (12, 44) + (18,) * len(mb))
    for row in ws.iter_rows(min_row=2, min_col=3):
        for c in row:
            c.number_format = '#,##0.00'
    colmes = {am: _col(3 + j) for j, am in enumerate(mb)}

    def expr(defi, am):
        partes = []
        for sg, fac, cta in TOKEN.findall(str(defi).replace(' ', '')):
            ref = f"Balance!{colmes[am]}{fila_cta[int(cta)]}"
            partes.append(f"{'-' if sg == '-' else '+'}{fac.replace(',', '.') + '*' if fac else ''}{ref}")
        return ''.join(partes).lstrip('+')

    def ant(am):
        y, m = divmod(am, 100)
        return _am_de(y - 1, 12) if m == 1 else am - 1

    ref_var = {}                                     # (hoja, columna) -> {am: 'Hoja!Celda'}
    ws = wb.create_sheet('Var_Balance')
    ws.append(['HOJA', 'VARIABLE', 'FUENTE', 'DEFINICION'] + [_etm(am) for am in meses])
    for i, v in enumerate(vb):
        r = i + 2
        fu = str(v['FUENTE']).strip().upper()
        fila = [v['HOJA'], v['COLUMNA'], fu, str(v['DEFINICION'])]
        for j, am in enumerate(meses):
            if not bal.get(am):
                fila.append(None)
                continue
            e = expr(v['DEFINICION'], am)
            if fu == 'BALANCE_PROMEDIO':
                ven = ventana(am, nprom)
                if not all(bal.get(a) for a in ven):
                    fila.append(None)
                    continue
                d_ = str(v['DEFINICION']).replace(' ', '')
                fila.append(f'=AVERAGE(Balance!{colmes[ven[0]]}{fila_cta[int(d_)]}:{colmes[ven[-1]]}{fila_cta[int(d_)]})' if d_.isdigit()
                            else '=(' + '+'.join(f'({expr(v["DEFINICION"], a)})' for a in ven) + f')/{nprom}')
            elif fu == 'BALANCE_MES' and am % 100 != 1:
                fila.append(f'=({e})-({expr(v["DEFINICION"], ant(am))})' if bal.get(ant(am)) else None)
            else:
                fila.append(f'={e}')
            ref_var.setdefault((v['HOJA'], v['COLUMNA']), {})[am] = f"Var_Balance!{_col(5 + j)}{r}"
        ws.append(fila)
    encab(ws)
    anchos(ws, (9, 46, 15, 32) + (18,) * len(meses))
    for row in ws.iter_rows(min_row=2, min_col=5):
        for c in row:
            c.number_format = '#,##0.00'

    # ---- Cartera
    if vc:
        ws = wb.create_sheet('Cartera_Datos')
        ws.append(['MES', 'CREDITO', 'LINEA', 'CATEGORIA', 'CALIF_OBJETIVA', 'SALDO_CAPITAL'])
        for am in meses:
            d = car.get(am)
            if d is None:
                continue
            for cr, li, ca, cl, sa in zip(d['CRED'], d['LINEA'], d['CAT'], d['CALIFICACION_OBJETIVA'], d['SALDO_CAPITAL']):
                ws.append([am, cr, li, ca, str(cl).strip().upper()[:1] if pd.notna(cl) else '', float(sa)])
        nc = ws.max_row
        encab(ws)
        anchos(ws, (9, 14, 40, 14, 10, 18))
        for (c,) in ws.iter_rows(min_row=2, min_col=6, max_col=6):
            c.number_format = '#,##0.00'
        R = {k: f'Cartera_Datos!${k}$2:${k}${nc}' for k in 'ACDEF'}
        R['L'] = f'Cartera_Datos!$C$2:$C${nc}'
        # transición
        ws = wb.create_sheet('Transicion_Datos')
        ws.append(['MES_ACTUAL', 'CREDITO', 'CALIF_ANTERIOR', 'CALIF_ACTUAL', 'SALDO_ANTERIOR', 'SALDO_ACTUAL', 'N_ANTERIOR', 'N_ACTUAL', 'MOVIMIENTO'])
        r = 2
        for am in meses:
            if car.get(am) is None or car.get(ant(am)) is None:
                continue
            p_, a_ = car[ant(am)], car[am]
            p_ = p_[(p_['SALDO_CAPITAL'] > 0) & p_['CAL'].notna()].groupby('CRED').agg(a=('CAL', 'first'), s=('SALDO_CAPITAL', 'sum'))
            a_ = a_[(a_['SALDO_CAPITAL'] > 0) & a_['CAL'].notna()].groupby('CRED').agg(x=('CAL', 'first'), s=('SALDO_CAPITAL', 'sum'))
            j = p_.join(a_, how='outer', lsuffix='_p', rsuffix='_a')
            for cr, q in j.iterrows():
                ws.append([am, cr, CAL[int(q['a'])] if pd.notna(q['a']) else '', CAL[int(q['x'])] if pd.notna(q['x']) else '',
                           float(q['s_p']) if pd.notna(q['s_p']) else 0.0, float(q['s_a']) if pd.notna(q['s_a']) else 0.0,
                           f'=IF(C{r}="","",CODE(C{r})-64)', f'=IF(D{r}="","",CODE(D{r})-64)',
                           f'=IF(C{r}="","Nuevo",IF(D{r}="",IF(G{r}>=2,"Cancelado B-E","Cancelado A"),IF(H{r}>G{r},"Deterioro",IF(H{r}<G{r},"Recuperación","Se mantiene"))))'])
                r += 1
        nt = ws.max_row
        encab(ws)
        anchos(ws, (11, 14, 11, 11, 18, 18, 10, 10, 16))
        for row in ws.iter_rows(min_row=2, min_col=5, max_col=6):
            for c in row:
                c.number_format = '#,##0.00'
        T = {k: f'Transicion_Datos!${k}$2:${k}${nt}' for k in 'AEFI'}
        ws = wb.create_sheet('Var_Cartera')
        ws.append(['HOJA', 'VARIABLE', 'DEFINICION'] + [_etm(am) for am in meses])
        for i, v in enumerate(vc):
            rr = i + 2
            k, _, arg = str(v['DEFINICION']).strip().partition(':')
            k = k.strip().upper()
            fila = [v['HOJA'], v['COLUMNA'], str(v['DEFINICION'])]
            for jj, am in enumerate(meses):
                if car.get(am) is None:
                    fila.append(None)
                    continue
                base = f"SUMIFS({R['F']},{R['A']},{am}"
                if k == 'TOTAL':
                    f = f'={base})'
                elif k == 'CATEGORIAS':
                    f = '=' + '+'.join(f'{base},{R["D"]},"{x.strip()}")' for x in arg.split(';') if x.strip())
                elif k == 'MORA_OBJETIVA':
                    f = '=' + '+'.join(f'{base},{R["E"]},"{x}")' for x in 'BCDE')
                elif k == 'LINEA':
                    f = f'={base},{R["L"]},"{arg.strip()}")'
                elif k == 'LINEA_A':
                    f = f'={base},{R["L"]},"{arg.strip()}",{R["E"]},"A")'
                elif k == 'TRANSICION_RECUPERACION':
                    f = (f'=SUMIFS({T["F"]},{T["A"]},{am},{T["I"]},"Recuperación")+SUMIFS({T["E"]},{T["A"]},{am},{T["I"]},"Cancelado B-E")'
                         if car.get(ant(am)) is not None else None)
                elif k == 'TRANSICION_DETERIORO':
                    f = f'=SUMIFS({T["F"]},{T["A"]},{am},{T["I"]},"Deterioro")' if car.get(ant(am)) is not None else None
                else:
                    f = None
                fila.append(f)
                if f:
                    ref_var.setdefault((v['HOJA'], v['COLUMNA']), {})[am] = f"Var_Cartera!{_col(4 + jj)}{rr}"
            ws.append(fila)
        encab(ws)
        anchos(ws, (9, 46, 40) + (18,) * len(meses))
        for row in ws.iter_rows(min_row=2, min_col=4):
            for c in row:
                c.number_format = '#,##0.00'

    # ---- Asociados
    if va:
        ws = wb.create_sheet('Asociados_Datos')
        ws.append(['MES', 'CEDULA', 'ESTABA_EN_DIC_ANTERIOR'])
        rango = {}
        r = 2
        for am in meses:
            s = aso.get(am)
            if s is None:
                continue
            r0 = r
            for ced in sorted(s):
                ws.append([am, ced])
                r += 1
            rango[am] = (r0, r - 1)
        for am, (r0, r1) in rango.items():
            dic = _am_de(am // 100 - 1, 12)
            if am % 100 == 12 and dic not in rango:
                continue
            if dic in rango:
                d0, d1 = rango[dic]
                for rr in range(r0, r1 + 1):
                    ws.cell(rr, 3).value = f'=IF(ISNUMBER(MATCH(B{rr},$B${d0}:$B${d1},0)),1,0)'
        encab(ws)
        anchos(ws, (9, 16, 14))
        wd = wb.create_sheet('Asociados_Dic')
        dics = sorted({_am_de(am // 100 - 1, 12) for am in rango if _am_de(am // 100 - 1, 12) in rango})
        meses_a = [am for am in rango if _am_de(am // 100 - 1, 12) in rango]
        wd.append(['DIC_BASE', 'CEDULA'] + [f'SIGUE EN {_etm(am)}' for am in meses_a])
        fila_dic = {}
        r = 2
        for dic in dics:
            d0, d1 = rango[dic]
            fila_dic[dic] = (r, r + d1 - d0)
            for k_ in range(d0, d1 + 1):
                fila = [dic, f'=Asociados_Datos!B{k_}']
                for am in meses_a:
                    if _am_de(am // 100 - 1, 12) == dic:
                        m0, m1 = rango[am]
                        fila.append(f'=IF(ISNUMBER(MATCH(B{r},Asociados_Datos!$B${m0}:$B${m1},0)),1,0)')
                    else:
                        fila.append(None)
                wd.append(fila)
                r += 1
        encab(wd)
        anchos(wd, (10, 16) + (12,) * len(meses_a))
        ws = wb.create_sheet('Var_Asociados')
        ws.append(['HOJA', 'VARIABLE', 'DEFINICION'] + [_etm(am) for am in meses])
        for i, v in enumerate(va):
            rr = i + 2
            k = str(v['DEFINICION']).strip().upper()
            fila = [v['HOJA'], v['COLUMNA'], k]
            for jj, am in enumerate(meses):
                f = None
                if am in rango:
                    r0, r1 = rango[am]
                    dic = _am_de(am // 100 - 1, 12)
                    tot = f'COUNTA(Asociados_Datos!$B${r0}:$B${r1})'
                    man = f'SUM(Asociados_Datos!$C${r0}:$C${r1})'
                    if k == 'TOTAL':
                        f = '=' + tot
                    elif dic in rango and k == 'BASE_DIC':
                        f = f'=COUNTA(Asociados_Datos!$B${rango[dic][0]}:$B${rango[dic][1]})'
                    elif dic in rango and k == 'SE_MANTIENEN':
                        f = '=' + man
                    elif dic in rango and k == 'NUEVOS':
                        f = f'={tot}-{man}'
                    elif dic in rango and k == 'RETIROS':
                        cm = _col(3 + meses_a.index(am))
                        a0, a1 = fila_dic[dic]
                        f = f'=COUNTIF(Asociados_Dic!${cm}${a0}:${cm}${a1},0)'
                fila.append(f)
                if f:
                    ref_var.setdefault((v['HOJA'], v['COLUMNA']), {})[am] = f"Var_Asociados!{_col(4 + jj)}{rr}"
            ws.append(fila)
        encab(ws)
        anchos(ws, (9, 46, 16) + (12,) * len(meses))

    # ---- Resumen y comparación
    ws = wb.create_sheet('Resumen', 1)
    ws.append(['HOJA', 'VARIABLE', 'FUENTE', 'DEFINICION'] + [_etm(am) for am in meses])
    for v in variables:
        refs = ref_var.get((v['HOJA'], v['COLUMNA']), {})
        ws.append([v['HOJA'], v['COLUMNA'], v['FUENTE'], str(v['DEFINICION'])] + [f'={refs[am]}' if am in refs else None for am in meses])
    encab(ws)
    anchos(ws, (9, 46, 15, 34) + (18,) * len(meses))
    for row in ws.iter_rows(min_row=2, min_col=5):
        for c in row:
            c.number_format = '#,##0.00'
    ws = wb.create_sheet('Comparacion', 2)
    ws.append(['HOJA', 'VARIABLE', 'AÑO', 'MES', 'CELDA DEL TABLERO', 'DIGITADO EN EL TABLERO', 'CALCULADO (FÓRMULA)', 'DIFERENCIA', 'DIF_RELATIVA', 'ACCION DEL EXTRACTOR'])
    r = 2
    for x in informe:
        hoja, col, y, mes, dig, _calc, _dif, _rel, acc, celda = x
        am = _am_de(int(y), MESES.index(mes) + 1)
        ref = ref_var.get((hoja, col), {}).get(am)
        ws.append([hoja, col, y, mes, celda, dig if isinstance(dig, (int, float)) else None, f'={ref}' if ref else None,
                   f'=IF(F{r}="","",G{r}-F{r})', f'=IF(OR(F{r}="",F{r}=0),"",ABS(H{r})/ABS(F{r}))', acc])
        r += 1
    encab(ws)
    anchos(ws, (9, 46, 7, 12, 10, 20, 20, 18, 12, 30))
    for row in ws.iter_rows(min_row=2, min_col=6, max_col=9):
        for c in row:
            c.number_format = '0.000%' if c.column == 9 else '#,##0.00'
    ws.auto_filter.ref = ws.dimensions
    for n in ('Resumen', 'Var_Balance', 'Var_Cartera', 'Var_Asociados', 'Comparacion'):
        if n in wb.sheetnames:
            wb[n].sheet_properties.tabColor = '0B5D3B'
    if os.path.exists(os.path.join(BASE, '~$' + os.path.basename(AUDITORIA))):
        print(f'   ! {os.path.basename(AUDITORIA)} está abierto en Excel: no se actualizó (ciérrelo y vuelva a ejecutar para regenerarlo)')
        return
    try:
        wb.save(AUDITORIA)
    except PermissionError:
        print(f'   ! No se pudo guardar {os.path.basename(AUDITORIA)} (¿abierto en Excel?): no se actualizó')
        return
    # Excel calcula y guarda los valores para que el libro se vea completo al abrirlo; se comprueba contra el cálculo de Python
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        xl = win32com.client.DispatchEx('Excel.Application')
        try:
            xl.Visible = False
            xl.DisplayAlerts = False
            w = xl.Workbooks.Open(AUDITORIA, UpdateLinks=0)
            xl.CalculateFull()
            malos = 0
            hr = w.Worksheets('Resumen')
            for i, v in enumerate(variables):
                for j, am in enumerate(meses):
                    py = res.get((v['HOJA'], v['COLUMNA']), {}).get(am)
                    xv = hr.Cells(i + 2, 5 + j).Value
                    if py is not None and (xv is None or abs(float(xv) - py) > 0.05):
                        malos += 1
            w.Save()
            w.Close(False)
        finally:
            xl.Quit()
            pythoncom.CoUninitialize()
        print(f'Auditoría: {AUDITORIA} · fórmulas de Excel frente al cálculo de Python: '
              + ('todas coinciden' if not malos else f'{malos} diferencias (revisar)'))
    except Exception as e:                                   # noqa: BLE001
        print(f'Auditoría: {AUDITORIA} (sin recalcular en Excel: {e}; Excel lo calcula al abrirlo)')


# ============================================================================================================
# 6. MAIN
# ============================================================================================================
def main():
    print(f'Extracción de indicadores del Plan Estratégico · v{VERSION}')
    gen, variables = leer_parametros()
    modo = str(gen.get('MODO', 'VACIOS')).strip().upper()
    for a, m in (('--comparar', 'COMPARAR'), ('--vacios', 'VACIOS'), ('--reemplazar', 'REEMPLAZAR')):
        if a in sys.argv:
            modo = m
    if modo not in ('COMPARAR', 'VACIOS', 'REEMPLAZAR'):
        raise ErrorExtraccion(f'MODO no válido: {modo} (COMPARAR, VACIOS o REEMPLAZAR)')
    tol = float(gen.get('TOLERANCIA_RELATIVA', 0.0001) or 0)
    solo = None
    if '--hojas' in sys.argv:
        k = sys.argv.index('--hojas')
        solo = {x.strip().upper() for x in (sys.argv[k + 1] if k + 1 < len(sys.argv) else '').split(',') if x.strip()}
        if not solo:
            raise ErrorExtraccion('Después de --hojas escriba las hojas separadas por coma, p. ej. --hojas IND-13,IND-06')
    print(f'Modo: {modo}' + (f' · solo hojas {", ".join(sorted(solo))}' if solo else ''))
    res, datos = calcular(gen, variables)
    informe, escribir = preparar(res, modo, tol, variables)
    if solo:                                                  # el informe es completo; solo se escriben las hojas pedidas
        escribir = [x for x in escribir if x[0].upper() in solo]
    guardar_informe(informe, modo, escribir)
    generar_auditoria(variables, datos, res, informe)
    print(f'Informe: {INFORME}')
    if modo != 'COMPARAR' and escribir:
        escribir_excel(escribir)
        print(f'Se escribieron {len(escribir)} celdas en {os.path.basename(TABLERO)}; ahora puede ejecutar Plan_Estrategico_Dashboard.py')
    else:
        print('No se escribió nada en el tablero.' if modo == 'COMPARAR' else 'No había celdas por escribir.')


if __name__ == '__main__':
    try:
        main()
    except ErrorExtraccion as e:
        print(f'\nERROR: {e}')
        sys.exit(1)
