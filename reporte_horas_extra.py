"""
Generador de Reporte de Horas Extra
Lee el reporte de Varias Obras y los reportes de obra (.xlsm) y genera un Excel
con una hoja Resumen y una hoja por reporte.

Columnas de cada hoja: Trabajador, Ajustero, Supervisor/Tecnico, Oficio, Sal. Dia,
Coste/Hora, horas extra por dia (una col por dia), Dias quincena trabajados (reales),
Domingos / Festivos (dias pagados doble), Coste dias / ajuste (ordinario, de la
plantilla), Total H. Extra, % H. Extra (= H. Extra / (dias + H. Extra)),
Coste H. Extra, Dias atrasados y Horas extras atrasadas (cantidades a mano),
Otros (ATRASOS/OTROS de la plantilla), Total atrasos,
Coste total quincena (= dias + H. Extra + atrasos) y comentarios.
"""

import sys, re, json, os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.comments import Comment
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas", "openpyxl"])
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.comments import Comment


# Arrastrar y soltar archivos (opcional: si tkinterdnd2 no esta disponible el
# programa funciona igual con los botones de "Añadir")
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
except ImportError:
    try:
        if getattr(sys, 'frozen', False):
            raise ImportError("ejecutable sin tkinterdnd2")
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "tkinterdnd2"])
        from tkinterdnd2 import TkinterDnD, DND_FILES
    except Exception:
        TkinterDnD, DND_FILES = None, None
except Exception:
    TkinterDnD, DND_FILES = None, None


# ---------------------------------------------------------------------------
# Configuración de umbrales (editable desde la interfaz, persistida en disco)
# ---------------------------------------------------------------------------

CONFIG_HORAS_DIA_ALERTA  = 5      # celda de día con más de N horas -> rojo/negrita
CONFIG_PORCENTAJE_MINIMO = 0.20   # % mínimo H. Extra sobre coste total -> alerta fila

_CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".reporte_horas_extra_config.json")

def cargar_configuracion():
    """Lee la configuración del disco si existe."""
    global CONFIG_HORAS_DIA_ALERTA, CONFIG_PORCENTAJE_MINIMO
    try:
        with open(_CONFIG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        CONFIG_HORAS_DIA_ALERTA  = int(data.get("horas_dia", CONFIG_HORAS_DIA_ALERTA))
        CONFIG_PORCENTAJE_MINIMO = float(data.get("porcentaje", CONFIG_PORCENTAJE_MINIMO))
    except Exception:
        pass

def guardar_configuracion():
    """Guarda la configuración actual en disco."""
    try:
        with open(_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump({"horas_dia":  CONFIG_HORAS_DIA_ALERTA,
                       "porcentaje": CONFIG_PORCENTAJE_MINIMO}, f)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Lectura de archivos
# ---------------------------------------------------------------------------

def _num(v):
    """Convierte a float; vacio / texto / NaN -> 0."""
    v = pd.to_numeric(v, errors='coerce')
    return float(v) if pd.notna(v) else 0.0

def _reparar(texto):
    """Corrige textos mal codificados del programa de nomina ('AlbaÃ±il' -> 'Albañil')."""
    if 'Ã' in texto or 'Â' in texto:
        try:
            return texto.encode('latin-1').decode('utf-8')
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return texto

def _txt(v):
    return _reparar(str(v).strip()) if pd.notna(v) else ''

def _norm(col):
    """Normaliza cabeceras: quita saltos de linea, _x000D_ y espacios dobles."""
    return ' '.join(str(col).replace('_x000D_', ' ').split()).upper()

def _fecha_key(fecha):
    """Clave de orden cronologico para fechas 'dd-mm-yyyy'."""
    try:
        return datetime.strptime(fecha, '%d-%m-%Y')
    except ValueError:
        return datetime.max


DIAS_SEMANA = ('Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom')
MESES = ('Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto',
         'Septiembre', 'Octubre', 'Noviembre', 'Diciembre')

# Festivos de fecha fija en Republica Dominicana (dia, mes). Los festivos que
# la ley traslada a lunes no se incluyen: se detectan por los dias marcados
# con 2 en los reportes de obra.
FESTIVOS_FIJOS = {(1, 1), (21, 1), (27, 2), (16, 8), (24, 9), (25, 12)}


def es_domingo_o_festivo(fecha, festivos=()):
    d = _fecha_key(fecha)
    if d == datetime.max:
        return False
    return d.weekday() == 6 or (d.day, d.month) in FESTIVOS_FIJOS or fecha in festivos


def etiqueta_dia(fecha):
    """'18-09-2026' -> 'Vie 18/09'"""
    d = _fecha_key(fecha)
    if d == datetime.max:
        return fecha
    return f"{DIAS_SEMANA[d.weekday()]} {d.day:02d}/{d.month:02d}"


def periodo_quincena(fechas):
    """
    A partir de las fechas de los reportes devuelve (nombre, desde, hasta),
    p. ej. ('2Q Septiembre 2026', '18/09/2026', '01/10/2026'). La quincena es
    del mes en que cae su dia central: 1Q si ese dia es <= 15, 2Q si no.
    """
    ds = sorted(d for d in (_fecha_key(f) for f in fechas) if d != datetime.max)
    if not ds:
        return None
    centro = ds[0] + (ds[-1] - ds[0]) / 2
    q = 1 if centro.day <= 15 else 2
    nombre = f"{q}Q {MESES[centro.month - 1]} {centro.year}"
    return nombre, ds[0].strftime('%d/%m/%Y'), ds[-1].strftime('%d/%m/%Y')


def _worker(nombre, ajustero, tecnico, supervisor, oficio, obra, sal_dia, coste_hora,
            dias_trab, dias_dobles, coste_dias, coste_total, atrasos, obs, days,
            gratificacion=0.0):
    total_e = sum(days.values())
    return {
        'TRABAJADOR':     nombre,
        'AJUSTERO':       ajustero,
        'TECNICO':        tecnico,
        'SUPERVISOR':     supervisor,
        'OFICIO':         oficio,
        'OBRA':           obra,
        'SAL. DIA':       sal_dia,
        'GRATIF':         gratificacion or None,
        'COSTE HORA':     coste_hora,
        'DIAS':           dias_trab,
        'DIAS DOBLES':    dias_dobles or None,
        'COSTE DIAS':     coste_dias,
        'COSTE TOTAL':    coste_total,
        'ATRASOS':        atrasos,
        'COMENTARIOS':    obs,
        'TOTAL H. EXTRA': total_e,
        'COSTE H. EXTRA': total_e * coste_hora,
        '_days':          days,
    }


def _es_codigo_obra(v):
    try:
        return 10000 <= int(v) <= 99999
    except Exception:
        return False


def _read_formato_diario(path):
    """
    Formato del reporte de Varias Obras: fila 1 con las fechas, fila 2 con
    cabeceras y por cada dia un bloque ADMIN, AJUSTE, OBRA 1, EXTRA 1, OBRA 2, EXTRA 2.
    Devuelve (workers_todos, coste_dias_todos, info) con info = {'fechas', 'festivos'}.
    """
    raw = pd.read_excel(path, header=None)
    date_row, header_row = raw.iloc[0], raw.iloc[1]

    cols = {}
    for i, h in enumerate(header_row):
        if pd.notna(h):
            cols.setdefault(_norm(h), i)

    day_blocks = []
    last_date  = None
    i = 0
    while i < len(header_row):
        m = re.search(r'(\d{2}-\d{2}-\d{4})', _txt(date_row.iloc[i]))
        if m:
            last_date = m.group(1)
        if last_date and _norm(header_row.iloc[i]) == 'ADMIN':
            day_blocks.append({
                'fecha':  last_date,
                'admin':  i,     'ajuste': i + 1,
                'obra1':  i + 2, 'extra1': i + 3,
                'obra2':  i + 4, 'extra2': i + 5,
            })
            i += 6
        else:
            i += 1

    def get(row, name):
        c = cols.get(name)
        return row.iloc[c] if c is not None else None

    workers = []
    for r in range(2, len(raw)):
        row    = raw.iloc[r]
        nombre = _txt(get(row, 'OPERARIOS'))
        if not nombre or nombre.upper() in ('TOTAL', 'TOTALES'):
            continue

        day_hours  = {}   # fecha -> horas extra del dia
        obra_hours = {}   # codigo_obra -> horas (regulares + extras)
        obra_def   = get(row, 'OBRA')
        obra_def   = str(int(obra_def)) if _es_codigo_obra(obra_def) else None

        for b in day_blocks:
            dias_reg = _num(row.iloc[b['admin']]) + _num(row.iloc[b['ajuste']])
            if dias_reg > 0 and obra_def:
                obra_hours[obra_def] = obra_hours.get(obra_def, 0) + dias_reg * 8
            for oc, ec in ((b['obra1'], b['extra1']), (b['obra2'], b['extra2'])):
                extra = _num(row.iloc[ec])
                if extra <= 0:
                    continue
                day_hours[b['fecha']] = day_hours.get(b['fecha'], 0) + extra
                if _es_codigo_obra(row.iloc[oc]):
                    cod = str(int(row.iloc[oc]))
                    obra_hours[cod] = obra_hours.get(cod, 0) + extra

        top = sorted(obra_hours.items(), key=lambda x: -x[1])[:2]

        # COSTE ADMIN incluye las horas extra y el recargo de los dias dobles
        # (festivos/domingos), que no se marcan dia a dia: se deducen del importe.
        sal_dia     = _num(get(row, 'SAL. DIA'))
        coste_hora  = _num(get(row, 'SAL. HORA'))
        dias_trab   = _num(get(row, 'TOTAL DIAS TRABAJADOS'))
        coste_admin = _num(get(row, 'COSTE ADMIN')) - sum(day_hours.values()) * coste_hora
        dobles      = (coste_admin - dias_trab * sal_dia) / sal_dia if sal_dia else 0
        dobles      = int(round(dobles)) if dobles > 0 and abs(dobles - round(dobles)) < 0.01 else 0

        workers.append(_worker(
            nombre      = nombre,
            ajustero    = _txt(get(row, 'AJUSTERO')).rstrip('*').strip(),
            tecnico     = _txt(get(row, 'TECNICO OBRA')),
            supervisor  = '',
            oficio      = _txt(get(row, 'OFICIO')),
            obra        = '/'.join(cod for cod, _ in top),
            sal_dia     = sal_dia,
            coste_hora  = coste_hora,
            dias_trab   = dias_trab,
            dias_dobles = dobles,
            coste_dias  = coste_admin + _num(get(row, 'COSTE AJUSTE')),
            coste_total = _num(get(row, 'IMPORTE TOTAL')),
            atrasos     = _num(get(row, 'ATRASOS/OTROS')),
            obs         = _txt(get(row, 'OBSERVACIONES')),
            days        = day_hours,
        ))

    info = {'fechas': [b['fecha'] for b in day_blocks], 'festivos': set()}
    return workers, sum(w['COSTE DIAS'] for w in workers), info


def read_varias_obras(path):
    """Devuelve: (nombre_hoja, workers, coste_dias_obra, total_trabajadores, info)"""
    todos, total_obra, info = _read_formato_diario(path)
    return ('Varias Obras', todos, float(total_obra), len(todos), info)


def read_obra_xlsm(path, sheet_name):
    """
    Reporte de obra (.xlsm, hoja 'Reporte', cabecera en la fila 4).
    Si el fichero no tiene hoja 'Reporte' se lee con el formato de Varias Obras.
    Devuelve: (nombre_hoja, workers, coste_dias_obra, total_trabajadores, info)
    info = {'fechas': fechas del reporte, 'festivos': fechas con algun dia doble}
    """
    if 'Reporte' not in pd.ExcelFile(path).sheet_names:
        todos, total_obra, info = _read_formato_diario(path)
        for w in todos:
            w['OBRA'] = ''
        return (sheet_name, todos, float(total_obra), len(todos), info)

    df = pd.read_excel(path, header=3, sheet_name='Reporte')
    df.columns = [_norm(c) for c in df.columns]
    df = df.dropna(subset=['OPERARIOS']).copy()
    df = df[~df['OPERARIOS'].astype(str).str.strip().str.upper().str.startswith('TOTAL')]

    # Columnas de horas extra por dia: "EXTRA 18/09/2026 (HORAS)"
    date_per_col = {}
    for col in df.columns:
        if col.startswith('EXTRA'):
            m = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', col)
            if m:
                d, mo, y = m.groups()
                date_per_col[col] = f"{int(d):02d}-{int(mo):02d}-{y}"

    # Columnas de dias por fecha: "ADMIN 18/09/2026 (DIAS)" / "AJUSTE ..."
    # Un dia marcado con 2 (festivo/domingo pagado doble) es un solo dia real.
    dias_per_col = {}
    for col in df.columns:
        if col.startswith('ADMIN') or col.startswith('AJUSTE'):
            m = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', col)
            if m:
                d, mo, y = m.groups()
                dias_per_col[col] = f"{int(d):02d}-{int(mo):02d}-{y}"
    festivos = set()

    def get(row, *names):
        for n in names:
            if n in row.index:
                return row[n]
        return None

    workers = []
    for _, row in df.iterrows():
        day_hours = {}
        for col, fecha in date_per_col.items():
            v = _num(row[col])
            if v > 0:
                day_hours[fecha] = day_hours.get(fecha, 0) + v

        dias_fecha = {}
        for col, fecha in dias_per_col.items():
            dias_fecha[fecha] = dias_fecha.get(fecha, 0) + _num(row[col])
        dias_trab = sum(1 for v in dias_fecha.values() if v > 0)
        dobles    = sum(v - 1 for v in dias_fecha.values() if v > 1)
        festivos.update(f for f, v in dias_fecha.items() if v > 1)

        sal_dia    = _num(get(row, 'SALARIO X DIA'))
        coste_hora = _num(get(row, 'COSTE HORAS EXTRA')) or sal_dia / 8
        atrasos    = _num(get(row, 'ATRASOS/OTROS'))
        if 'COSTE TOTAL' in row.index:
            coste_total = _num(row['COSTE TOTAL'])
        else:
            coste_total = _num(get(row, 'TOTAL A PAGAR'))
        if 'COSTE DIAS' in row.index:
            coste_dias = _num(row['COSTE DIAS'])
        else:
            coste_dias = coste_total - atrasos - sum(day_hours.values()) * coste_hora

        workers.append(_worker(
            nombre      = _txt(row['OPERARIOS']),
            ajustero    = _txt(get(row, 'AJUSTERO')),
            tecnico     = '',
            supervisor  = _txt(get(row, 'SUPERVISOR')),
            oficio      = _txt(get(row, 'OFICIO')),
            obra        = '',
            sal_dia     = sal_dia,
            gratificacion = _num(get(row, 'GRATIFICACIONES')),
            coste_hora  = coste_hora,
            dias_trab   = dias_trab,
            dias_dobles = dobles,
            coste_dias  = coste_dias,
            coste_total = coste_total,
            atrasos     = atrasos,
            obs         = _txt(get(row, 'OBSERVACIONES')),
            days        = day_hours,
        ))

    total_obra = sum(w['COSTE DIAS'] for w in workers)
    info = {'fechas': sorted(set(dias_per_col.values()) | set(date_per_col.values()), key=_fecha_key),
            'festivos': festivos}
    return (sheet_name, workers, float(total_obra), len(workers), info)


def detect_obra_code(path):
    """Extrae el codigo de obra (5 digitos exactos) del nombre del fichero."""
    name = os.path.basename(str(path))
    nums = re.findall(r'(?<!\d)\d{5}(?!\d)', name)
    return nums[0] if nums else None


# ---------------------------------------------------------------------------
# Escritura del Excel
# ---------------------------------------------------------------------------

# Paleta azul — visualmente intuitiva, accesible para consulta rapida
HDR_BG   = "1F4E79"   # azul oscuro (cabecera)
TOTAL_BG = "E26B0A"   # naranja: 1er pago (ordinario de la quincena)
PAGO2_BG = "B4A7D6"   # lila: 2o pago (horas extra y atrasos, tras autorizarse)
TOTALQ_BG = "000000"  # negro: coste total de la quincena (1er + 2o pago)
DOM_HDR  = "A6A6A6"   # gris: cabecera de domingos y festivos
DOM_BG   = "EDEDED"   # gris claro: celdas de domingos y festivos
HDR_FG   = "FFFFFF"
DAY_BG   = "BDD7EE"   # azul claro para cols de dias
TOT_BG   = "9DC3E6"   # azul medio para fila total
ALERT_FG = "C00000"   # rojo intenso para alertas (texto)

FMT_MONEY = '#,##0.00'
FMT_ACCT  = '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-'
FMT_NUM   = 'General'   # horas y dias: enteros sin decimales, fracciones si las hay
FMT_PCT   = '0.0%'

def _brd_outer():
    """Borde exterior completo."""
    s = Side(style='thin', color='999999')
    return Border(left=s, right=s, top=s, bottom=s)

def _brd_row():
    """Solo linea inferior (separador de filas)."""
    b = Side(style='thin', color='CCCCCC')
    n = Side(style=None)
    return Border(bottom=b, left=n, right=n, top=n)

def _hdr_font():  return Font(name='Calibri', bold=True, color=HDR_FG, size=10)
def _hdr_fill():  return PatternFill('solid', start_color=HDR_BG)
def _total_hdr_fill(): return PatternFill('solid', start_color=TOTAL_BG)
def _day_fill():  return PatternFill('solid', start_color=DAY_BG)
def _tot_fill():  return PatternFill('solid', start_color=TOT_BG)
def _cen():       return Alignment(horizontal='center', vertical='center')
def _left():      return Alignment(horizontal='left',   vertical='center')


TEXT_COLS = ('TRABAJADOR', 'AJUSTERO', 'TECNICO', 'SUPERVISOR', 'OFICIO', 'OBRA', 'COMENTARIOS')

LABELS = {
    'TRABAJADOR':     'Trabajador',
    'AJUSTERO':       'Ajustero',
    'TECNICO':        'Técnico Obra',
    'SUPERVISOR':     'Supervisor',
    'OFICIO':         'Oficio',
    'OBRA':           'Obra',
    'SAL. DIA':       'Sal. Dia',
    'GRATIF':         'Gratificación',
    'COSTE HORA':     'Coste/Hora',
    'DIAS':           'Dias quincena trabajados',
    'DIAS DOBLES':    'Domingos / Festivos',
    'COSTE DIAS':     'Coste dias / ajuste',
    'COSTE H. EXTRA': 'Coste H. Extra',
    'TOTAL H. EXTRA': 'Total H. Extra',
    '% H. EXTRA':     '% H. Extra',
    'COSTE TOTAL':    'Coste total quincena',
    'DIAS ATR':       'Dias atrasados',
    'HORAS ATR':      'Horas extras atrasadas',
    'ATRASOS':        'Otros',
    'TOTAL ATR':      'Total atrasos',
    'A PAGAR':        'A pagar H.E. + Atrasos',
    'COMENTARIOS':    'comentarios',
}

# Proceso de pago de cada quincena:
#   1er pago (naranja): pago ordinario de la quincena -> Coste dias / ajuste.
#   2o pago (lila): cuando se autorizan, horas extra y atrasos -> A pagar H.E. + Atrasos.
#   Coste total quincena (negro) = 1er pago + 2o pago.
PAGO_HDR = {
    'COSTE DIAS':     TOTAL_BG,
    'COSTE H. EXTRA': PAGO2_BG,
    'TOTAL ATR':      PAGO2_BG,
    'A PAGAR':        PAGO2_BG,
    'COSTE TOTAL':    TOTALQ_BG,
}

# Nota en la cabecera de las columnas de pagos y totales
NOTAS = {
    'COSTE DIAS':     '1er PAGO · Pago ordinario de la quincena (días trabajados + ajuste), '
                      'copiado del reporte. Se paga en la propia quincena. '
                      'No incluye horas extra ni atrasos.',
    'COSTE H. EXTRA': 'Parte del 2º PAGO · Total H. Extra × Coste/Hora. '
                      'Se paga cuando se autorizan las horas extra.',
    'TOTAL ATR':      'Parte del 2º PAGO · Dias atrasados × Sal. Dia + '
                      'Horas extras atrasadas × Coste/Hora + Otros.',
    'A PAGAR':        '2º PAGO · Coste H. Extra + Total atrasos. Es lo que queda por pagar '
                      'al trabajador después del pago ordinario, una vez autorizado.',
    'COSTE TOTAL':    'COSTE TOTAL DE LA QUINCENA · 1er pago (Coste dias / ajuste) + '
                      '2º pago (A pagar H.E. + Atrasos). Es lo que se carga a la obra.',
}

RESUMEN_HEADERS = ['Obra', 'Trabajadores', 'Trab. con HE', 'Coste dias quincena',
                   'Coste H. Extra', 'Total H. Extra', '% s/Total', 'Coste atrasos',
                   'A pagar H.E. + Atrasos', 'Coste Total (Dias + HE + Atrasos)']

RESUMEN_HDR = {
    'Coste dias quincena':               TOTAL_BG,
    'Coste H. Extra':                    PAGO2_BG,
    'Coste atrasos':                     PAGO2_BG,
    'A pagar H.E. + Atrasos':            PAGO2_BG,
    'Coste Total (Dias + HE + Atrasos)': TOTALQ_BG,
}

NOTAS_RESUMEN = {
    'Coste dias quincena':  '1er PAGO · Pago ordinario (días + ajuste) de todas las personas '
                            'del reporte. Se paga en la propia quincena.',
    'Coste H. Extra':       'Parte del 2º PAGO · Suma de Coste H. Extra de la hoja de la obra.',
    'Coste atrasos':        'Parte del 2º PAGO · Suma de Total atrasos de la hoja de la obra.',
    'A pagar H.E. + Atrasos':
                            '2º PAGO · Coste H. Extra + Coste atrasos. Lo que queda por pagar '
                            'de la obra una vez autorizado.',
    'Coste Total (Dias + HE + Atrasos)':
                            'COSTE TOTAL DE LA QUINCENA · 1er pago (Coste dias quincena) + '
                            '2º pago (A pagar H.E. + Atrasos). Coste de mano de obra de la obra.',
}


def _nota(texto):
    c = Comment(texto, 'Reporte Horas Extra')
    c.width, c.height = 260, 90
    return c


COL_WIDTHS = {
    'TRABAJADOR': 30, 'AJUSTERO': 26, 'TECNICO': 18, 'SUPERVISOR': 20, 'OFICIO': 14,
    'OBRA': 14, 'SAL. DIA': 10, 'GRATIF': 13, 'COSTE HORA': 11, 'DIAS': 25, 'DIAS DOBLES': 20,
    'COSTE DIAS': 22,
    'COSTE H. EXTRA': 14, 'TOTAL H. EXTRA': 13, '% H. EXTRA': 12, 'COSTE TOTAL': 20,
    'DIAS ATR': 14, 'HORAS ATR': 21, 'ATRASOS': 16, 'TOTAL ATR': 16, 'A PAGAR': 22,
    'COMENTARIOS': 45,
}

# Columnas que reunen totales: cabecera en naranja para distinguirlas
TOTAL_HDR_COLS = tuple(PAGO_HDR)

# Columnas que llevan SUM en la fila de totales
SUM_COLS = ('DIAS', 'DIAS DOBLES', 'COSTE DIAS', 'COSTE H. EXTRA', 'TOTAL H. EXTRA',
            'COSTE TOTAL', 'DIAS ATR', 'HORAS ATR', 'ATRASOS', 'TOTAL ATR', 'A PAGAR')

# Columnas de importes en formato contable (vacias o cero se ven como "-")
ACCT_COLS = ('COSTE DIAS', 'ATRASOS', 'TOTAL ATR')

# Columnas de horas / dias (sin formato decimal)
NUM_COLS = ('DIAS', 'DIAS DOBLES', 'TOTAL H. EXTRA', 'DIAS ATR', 'HORAS ATR')


def _pct_he(w):
    """% de horas extra sobre lo generado en la quincena (dias + horas extra)."""
    base = w['COSTE DIAS'] + w['COSTE H. EXTRA']
    return w['COSTE H. EXTRA'] / base if base > 0 else 0


def write_sheet(wb, sheet_name, workers, info=None):
    """
    Escribe la hoja de un reporte. Devuelve la fila de totales (o None si no hay
    trabajadores) y el mapa columna -> letra, para enlazar desde el Resumen.
    """
    ws = wb.create_sheet(title=sheet_name[:31])

    if not workers:
        ws['A1'] = 'Sin trabajadores en el reporte'
        return None, {}

    # Todos los dias de la quincena (aunque nadie tenga horas extra ese dia)
    info     = info or {}
    all_days = sorted(set(info.get('fechas', [])) | {d for w in workers for d in w['_days']},
                      key=_fecha_key)
    festivos = {d for d in all_days if es_domingo_o_festivo(d, info.get('festivos', ()))}

    fixed = ['TRABAJADOR', 'AJUSTERO']
    for key in ('TECNICO', 'SUPERVISOR', 'OFICIO', 'OBRA'):
        if any(w.get(key) for w in workers):
            fixed.append(key)
    fixed.append('SAL. DIA')
    if any(w.get('GRATIF') for w in workers):
        fixed.append('GRATIF')          # gratificacion diaria: sube el Coste/Hora
    fixed.append('COSTE HORA')

    tail = ['DIAS']
    if any(w.get('DIAS DOBLES') for w in workers):
        tail.append('DIAS DOBLES')
    tail += ['COSTE DIAS', 'TOTAL H. EXTRA', '% H. EXTRA', 'COSTE H. EXTRA',
             'DIAS ATR', 'HORAS ATR', 'ATRASOS', 'TOTAL ATR', 'A PAGAR', 'COSTE TOTAL',
             'COMENTARIOS']

    all_cols = fixed + all_days + tail
    L = {col: get_column_letter(i) for i, col in enumerate(all_cols, 1)}

    # Umbral dinamico de horas extra (sobre trabajadores con horas extra):
    #   - al menos 2.3x la media Y al menos 25h por encima de la media
    con_he     = [w for w in workers if w['TOTAL H. EXTRA'] > 0]
    horas      = [w['TOTAL H. EXTRA'] for w in con_he]
    media      = sum(horas) / len(horas) if horas else 0
    umbral_h   = max(media * 2.3, media + 25)

    # Umbral dinamico de % H. Extra sobre coste total:
    #   - al menos 1.5x el % medio Y 10 puntos por encima, nunca bajo el minimo
    pcts       = [_pct_he(w) for w in con_he if w['COSTE DIAS'] + w['COSTE H. EXTRA'] > 0]
    media_pct  = sum(pcts) / len(pcts) if pcts else 0
    umbral_pct = max(media_pct * 1.5, media_pct + 0.10, CONFIG_PORCENTAJE_MINIMO)

    # ── Fila 1: cabecera ─────────────────────────────────────────────────────
    for ci, col in enumerate(all_cols, 1):
        if col in all_days:
            c = ws.cell(1, ci, etiqueta_dia(col))
            c.fill = PatternFill('solid', start_color=DOM_HDR) if col in festivos else _day_fill()
            c.font = Font(name='Calibri', bold=True, color=HDR_BG, size=9)
        else:
            c = ws.cell(1, ci, LABELS.get(col, col))
            if col in NOTAS:            # solo columnas de totales (naranja)
                c.comment = _nota(NOTAS[col])
            c.fill = PatternFill('solid', start_color=PAGO_HDR[col]) if col in PAGO_HDR else _hdr_fill()
            c.font = _hdr_font()
        c.alignment = _cen()
        c.border    = _brd_outer()
    ws.row_dimensions[1].height = 20

    # ── Filas de datos ───────────────────────────────────────────────────────
    first_day, last_day = (L[all_days[0]], L[all_days[-1]]) if all_days else (None, None)

    for r, w in enumerate(workers, 2):
        pct_extra   = _pct_he(w)
        es_alerta   = w['TOTAL H. EXTRA'] > umbral_h or (pct_extra > umbral_pct and pct_extra > 0)

        # Coste dias / ajuste y Otros (ATRASOS/OTROS) se trasladan tal cual de
        # la plantilla; Dias atrasados y Horas extras atrasadas son cantidades
        # (dias / horas) que se rellenan a mano y se valoran a Sal. Dia y Coste/Hora.
        # Coste total quincena = coste ordinario + coste total.
        formulas = {
            'COSTE H. EXTRA': f"={L['TOTAL H. EXTRA']}{r}*{L['COSTE HORA']}{r}",
            'TOTAL H. EXTRA': f"=SUM({first_day}{r}:{last_day}{r})" if all_days else 0,
            '% H. EXTRA':     f"=IFERROR({L['COSTE H. EXTRA']}{r}/({L['COSTE DIAS']}{r}+{L['COSTE H. EXTRA']}{r}),0)",
            'A PAGAR':        f"=N({L['COSTE H. EXTRA']}{r})+N({L['TOTAL ATR']}{r})",
            'COSTE TOTAL':    f"=N({L['COSTE DIAS']}{r})+N({L['A PAGAR']}{r})",
            'TOTAL ATR':      (f"=N({L['DIAS ATR']}{r})*{L['SAL. DIA']}{r}"
                               f"+N({L['HORAS ATR']}{r})*{L['COSTE HORA']}{r}"
                               f"+N({L['ATRASOS']}{r})"),
        }

        for ci, col in enumerate(all_cols, 1):
            if col in formulas:
                val = formulas[col]
            elif col in all_days:
                val = w['_days'].get(col)
            elif col in ('DIAS ATR', 'HORAS ATR'):
                val = None
            else:
                val = w.get(col)
            if val == '' or (isinstance(val, float) and val == 0.0 and col in all_days):
                val = None

            bold  = (col in ('TOTAL H. EXTRA', '% H. EXTRA') and es_alerta) or col in ('A PAGAR', 'COSTE TOTAL')
            color = ALERT_FG if es_alerta else '000000'
            if col in all_days and val is not None and val > CONFIG_HORAS_DIA_ALERTA:
                color, bold = ALERT_FG, True

            c = ws.cell(r, ci, val)
            c.font      = Font(name='Calibri', size=10, bold=bold, color=color)
            c.border    = _brd_row()
            c.alignment = _left() if col in TEXT_COLS else _cen()
            if col in festivos:
                c.fill = PatternFill('solid', start_color=DOM_BG)
            if col in ('SAL. DIA', 'GRATIF', 'COSTE HORA', 'COSTE H. EXTRA', 'A PAGAR', 'COSTE TOTAL'):
                c.number_format = FMT_MONEY
            elif col in ACCT_COLS:
                c.number_format = FMT_ACCT
            elif col in NUM_COLS or col in all_days:
                c.number_format = FMT_NUM
            elif col == '% H. EXTRA':
                c.number_format = FMT_PCT

    # ── Fila totales ─────────────────────────────────────────────────────────
    last_data = 1 + len(workers)
    tr        = last_data + 1
    for ci, col in enumerate(all_cols, 1):
        c = ws.cell(tr, ci)
        c.fill      = _tot_fill()
        c.font      = Font(name='Calibri', bold=True, size=10)
        c.border    = _brd_outer()
        c.alignment = _left() if ci == 1 else _cen()
        cl = L[col]
        if ci == 1:
            c.value = 'Total'
        elif col in all_days or col in SUM_COLS:
            c.value = f'=SUM({cl}2:{cl}{last_data})'
            if col in all_days or col in NUM_COLS:
                c.number_format = FMT_NUM
            elif col in ('COSTE H. EXTRA', 'A PAGAR', 'COSTE TOTAL'):
                c.number_format = FMT_MONEY
            else:
                c.number_format = FMT_ACCT
        elif col == '% H. EXTRA':
            c.value         = f"=IFERROR({L['COSTE H. EXTRA']}{tr}/({L['COSTE DIAS']}{tr}+{L['COSTE H. EXTRA']}{tr}),0)"
            c.number_format = FMT_PCT

    # ── Anchos, filtros y paneles ────────────────────────────────────────────
    for ci, col in enumerate(all_cols, 1):
        ws.column_dimensions[get_column_letter(ci)].width = COL_WIDTHS.get(col, 10)

    ws.auto_filter.ref = f"A1:{get_column_letter(len(all_cols))}{last_data}"
    ws.freeze_panes = 'A2'
    return tr, L


def _ref(sheet_title, cell):
    """Referencia a celda de otra hoja: 'Varias Obras'!Y45"""
    return "'" + sheet_title.replace("'", "''") + "'!" + cell


def periodo_de(sheets):
    """
    Quincena de un conjunto de reportes leidos: (nombre, desde, hasta) o None.
    Si los reportes no son todos de la misma quincena, se toma la que mas se
    repite (y, en empate, la mas reciente).
    """
    periodos = [p for p in (periodo_quincena(sh[4].get('fechas', [])) for sh in sheets) if p]
    if not periodos:
        return None
    return max(periodos, key=lambda p: (periodos.count(p),
                                        datetime.strptime(p[2], '%d/%m/%Y')))


def build_excel(sheets, output_path):
    """sheets es lista de (sheet_name, workers, coste_dias_obra, total_trabajadores, info)."""
    wb = Workbook()
    wb.remove(wb.active)
    periodo = periodo_de(sheets)

    # ── Hoja resumen ──────────────────────────────────────────────────────────
    # Titulo con la quincena y debajo la tabla:
    # Obra | Trabajadores | Trab. con HE | Coste dias quincena | Coste H. Extra
    #      | Total H. Extra | % s/Total | Coste atrasos | Coste Total
    # Coste dias quincena es el coste ordinario de TODOS los trabajadores del
    # reporte; el resto se enlaza con los totales de cada hoja, de forma que al
    # rellenar atrasos a mano el resumen se actualiza solo.
    ws_r = wb.create_sheet('Resumen', 0)

    headers = RESUMEN_HEADERS
    n_cols  = len(headers)
    last_cl = get_column_letter(n_cols)

    t = ws_r.cell(1, 1, 'Reporte de Horas Extra y Atrasos' + (f' · {periodo[0]}' if periodo else ''))
    t.font = Font(name='Calibri', bold=True, size=14, color=HDR_BG)
    ws_r.merge_cells(f'A1:{last_cl}1')
    ws_r.row_dimensions[1].height = 22

    HR = 3                      # fila de cabecera de la tabla
    for ci, label in enumerate(headers, 1):
        c = ws_r.cell(HR, ci, label)
        c.font      = _hdr_font()
        # Mismos colores de pago que en las hojas
        c.fill      = PatternFill('solid', start_color=RESUMEN_HDR[label]) if label in RESUMEN_HDR else _hdr_fill()
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border    = _brd_outer()
        if label in NOTAS_RESUMEN:      # solo columnas de pagos y totales
            c.comment = _nota(NOTAS_RESUMEN[label])
    ws_r.row_dimensions[HR].height = 30

    money_cols = (4, 5, 8, 9, 10)
    first = HR + 1
    for ri, (sname, workers, coste_dias_obra, total_trab, info) in enumerate(sheets, first):
        tr, L = write_sheet(wb, sname, workers, info)
        title = sname[:31]
        if tr:
            last = tr - 1
            con_he   = f"=COUNTIF({_ref(title, L['TOTAL H. EXTRA'] + '2')}:{L['TOTAL H. EXTRA']}{last},\">0\")"
            coste_he = '=' + _ref(title, f"{L['COSTE H. EXTRA']}{tr}")
            total_he = '=' + _ref(title, f"{L['TOTAL H. EXTRA']}{tr}")
            atrasos  = '=' + _ref(title, f"{L['TOTAL ATR']}{tr}")
        else:
            con_he = coste_he = total_he = atrasos = 0

        valores = [sname, total_trab, con_he, coste_dias_obra, coste_he, total_he,
                   f'=IFERROR(E{ri}/(D{ri}+E{ri}),0)', atrasos,
                   f'=E{ri}+H{ri}', f'=D{ri}+I{ri}']
        for ci, val in enumerate(valores, 1):
            c = ws_r.cell(ri, ci, val)
            c.font      = Font(name='Calibri', size=10, bold=(ci >= n_cols - 1))
            c.border    = _brd_row()
            c.alignment = _left() if ci == 1 else _cen()
            if ci in money_cols: c.number_format = FMT_MONEY
            if ci == 6:          c.number_format = FMT_NUM
            if ci == 7:          c.number_format = FMT_PCT

    tr = first + len(sheets)
    for ci in range(1, n_cols + 1):
        c = ws_r.cell(tr, ci)
        c.fill      = _tot_fill()
        c.font      = Font(name='Calibri', bold=True, size=10)
        c.border    = _brd_outer()
        c.alignment = _left() if ci == 1 else _cen()
        cl = get_column_letter(ci)
        if ci == 1:
            c.value = 'Total'
        elif ci == 7:
            # % global = coste horas extra / (coste dias + coste horas extra)
            c.value = f'=IFERROR(E{tr}/(D{tr}+E{tr}),0)'
            c.number_format = FMT_PCT
        else:
            c.value = f'=SUM({cl}{first}:{cl}{tr-1})'
            if ci in money_cols: c.number_format = FMT_MONEY
            if ci == 6:          c.number_format = FMT_NUM

    for col, width in zip('ABCDEFGHIJ', (22, 13, 13, 20, 16, 15, 12, 16, 22, 32)):
        ws_r.column_dimensions[col].width = width
    ws_r.auto_filter.ref = f"A{HR}:{last_cl}{tr-1}"
    ws_r.freeze_panes = f'A{HR + 1}'

    wb.save(output_path)


# ---------------------------------------------------------------------------
# Interfaz grafica
# ---------------------------------------------------------------------------




def resource_path(filename):
    """Devuelve la ruta al recurso, compatible con PyInstaller (--onefile)."""
    import os
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, filename)


EXCEL_EXT = ('.xlsx', '.xlsm', '.xls')

def es_reporte_varias(path):
    """Un reporte de Varias Obras lleva 'VARIAS' en el nombre del fichero."""
    return 'VARIAS' in os.path.basename(path).upper()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        # Carga la extension nativa de arrastrar y soltar; si falla, solo botones
        self._dnd = False
        if TkinterDnD:
            try:
                TkinterDnD._require(self)
                self._dnd = True
            except Exception:
                pass
        self.title("Reporte de Horas Extra")
        self.resizable(False, False)
        self.configure(bg="#f0f4f8")
        try:
            self.iconbitmap(resource_path("logo_programa_multi.ico"))
        except Exception:
            pass
        # Estado
        self._varias_path  = None  # str o None
        self._varias_row   = None  # widget de la fila
        self._obras        = []    # lista de (path, sheet_label, widget_frame)
        # UI
        self._build_ui()
        self.eval("tk::PlaceWindow . center")

    # ── Construcción de la UI ────────────────────────────────────────────────
    def _build_ui(self):
        tk.Label(self, text="Reporte de Horas Extra",
                 font=("Arial", 14, "bold"), bg="#1F4E79", fg="white",
                 pady=10, padx=20).grid(row=0, column=0, sticky="ew")

        main = tk.Frame(self, bg="#f0f4f8", padx=16, pady=12)
        main.grid(row=1, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)

        # ── Sección Varias Obras ─────────────────────────────────────────────
        sec_varias = tk.LabelFrame(main, text=" Reporte de Varias Obras ",
                                   font=("Arial", 10, "bold"),
                                   fg="#1F4E79", bg="#f0f4f8",
                                   padx=10, pady=8)
        sec_varias.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        sec_varias.columnconfigure(0, weight=1)

        self._frame_varias = tk.Frame(sec_varias, bg="#f0f4f8")
        self._frame_varias.grid(row=0, column=0, sticky="ew")
        self._frame_varias.columnconfigure(0, weight=1)

        self._lbl_empty_varias = tk.Label(self._frame_varias,
            text="Aún no se ha añadido el reporte de Varias Obras."
                 + (" Puedes arrastrarlo aquí." if self._dnd else ""),
            font=("Arial", 9, "italic"), bg="#f0f4f8", fg="#888")
        self._lbl_empty_varias.grid(row=0, column=0, pady=4, sticky="w")

        self._btn_add_varias = tk.Button(sec_varias,
                  text="+  Añadir reporte de Varias Obras",
                  font=("Arial", 9), relief="flat",
                  bg="#dce6f1", cursor="hand2",
                  command=self._add_varias)
        self._btn_add_varias.grid(row=1, column=0, sticky="w", pady=(8, 0))

        # ── Sección Obras adicionales ────────────────────────────────────────
        sec_obras = tk.LabelFrame(main, text=" Reportes de obra ",
                                  font=("Arial", 10, "bold"),
                                  fg="#1F4E79", bg="#f0f4f8",
                                  padx=10, pady=8)
        sec_obras.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        sec_obras.columnconfigure(0, weight=1)

        self._frame_obras = tk.Frame(sec_obras, bg="#f0f4f8")
        self._frame_obras.grid(row=0, column=0, sticky="ew")
        self._frame_obras.columnconfigure(0, weight=1)

        self._lbl_empty_obras = tk.Label(self._frame_obras,
            text="Aún no se ha añadido ningún reporte de obra."
                 + (" Puedes arrastrarlos aquí." if self._dnd else ""),
            font=("Arial", 9, "italic"), bg="#f0f4f8", fg="#888")
        self._lbl_empty_obras.grid(row=0, column=0, pady=4, sticky="w")

        tk.Button(sec_obras, text="+  Añadir reporte(s) de obra",
                  font=("Arial", 9), relief="flat",
                  bg="#dce6f1", cursor="hand2",
                  command=self._add_obras
                  ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        # ── Botones Generar + Ajustes ────────────────────────────────────────
        btn_row = tk.Frame(main, bg="#f0f4f8")
        btn_row.grid(row=2, column=0, pady=(4, 4))

        tk.Button(btn_row, text="▶  Generar informe", font=("Arial", 11, "bold"),
                  bg="#4CAF50", fg="white", activebackground="#388E3C",
                  relief="flat", cursor="hand2", padx=18, pady=8,
                  command=self._run
                  ).grid(row=0, column=0, padx=(0, 8))

        tk.Button(btn_row, text="⚙  Ajustes", font=("Arial", 10),
                  bg="#dce6f1", relief="flat", cursor="hand2",
                  padx=12, pady=8,
                  command=self._open_settings
                  ).grid(row=0, column=1)

        self.progress = ttk.Progressbar(main, mode="indeterminate", length=420)
        self.progress.grid(row=3, column=0, sticky="ew", pady=(0, 2))

        self.lbl_status = tk.Label(main, text="", font=("Arial", 9),
                                   bg="#f0f4f8", fg="#444")
        self.lbl_status.grid(row=4, column=0, pady=(0, 4))

        self._setup_dnd(sec_varias, sec_obras)

    # ── Helper: crear fila visual de reporte cargado ─────────────────────────
    def _make_row(self, parent, label_text, path, on_remove):
        row = tk.Frame(parent, bg="#f0f4f8")
        row.columnconfigure(2, weight=1)
        tk.Label(row, text="OK", font=("Arial", 9, "bold"),
                 bg="#4CAF50", fg="white", width=3, pady=2
                 ).grid(row=0, column=0, padx=(0, 8))
        tk.Label(row, text=label_text, font=("Arial", 10, "bold"),
                 bg="#f0f4f8", fg="#1F4E79", width=12, anchor="w"
                 ).grid(row=0, column=1, padx=(0, 6), sticky="w")
        short = str(path).split("/")[-1].split("\\")[-1]
        tk.Label(row, text=short, font=("Arial", 9),
                 bg="#f0f4f8", fg="#666", anchor="w"
                 ).grid(row=0, column=2, sticky="w")
        tk.Button(row, text="✕", font=("Arial", 9, "bold"),
                  relief="flat", bg="#f0f4f8", fg="#c00", cursor="hand2",
                  command=on_remove
                  ).grid(row=0, column=3, padx=(6, 0))
        self._register_drop(row, 'varias' if parent is self._frame_varias else 'obras')
        return row

    # ── Acciones: Varias Obras ───────────────────────────────────────────────
    def _add_varias(self, p=None):
        if p is None:
            p = filedialog.askopenfilename(
                title="Selecciona el reporte de Varias Obras",
                filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
        if not p:
            return
        if self._varias_path:
            self._remove_varias()
        self._varias_path = p
        if self._lbl_empty_varias.winfo_exists():
            self._lbl_empty_varias.grid_remove()
        self._varias_row = self._make_row(self._frame_varias,
                                          "Varias Obras", p,
                                          on_remove=self._remove_varias)
        self._varias_row.grid(row=1, column=0, sticky="ew", pady=2)
        self._btn_add_varias.config(state="disabled")
        self.update_idletasks()
        self.geometry("")

    def _remove_varias(self):
        if self._varias_row is not None:
            self._varias_row.destroy()
            self._varias_row = None
        self._varias_path = None
        self._lbl_empty_varias.grid()
        self._btn_add_varias.config(state="normal")
        self.update_idletasks()
        self.geometry("")

    # ── Acciones: Reportes de obra (multiselección) ──────────────────────────
    def _add_obras(self, paths=None):
        if paths is None:
            paths = filedialog.askopenfilenames(
                title="Selecciona uno o varios reportes de obra",
                filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
        if not paths:
            return
        if self._lbl_empty_obras.winfo_exists():
            self._lbl_empty_obras.grid_remove()
        for p in paths:
            # Evitar duplicados (si el usuario añade el mismo archivo dos veces)
            if any(existing == p for existing, _, _ in self._obras):
                continue
            cod = detect_obra_code(p)
            sheet_label = cod if cod else f"Reporte {len(self._obras) + 1}"
            row = self._make_row(self._frame_obras, sheet_label, p,
                                 on_remove=lambda pp=p: self._remove_obra(pp))
            row.grid(row=len(self._obras), column=0, sticky="ew", pady=2)
            self._obras.append((p, sheet_label, row))
        self.update_idletasks()
        self.geometry("")

    # ── Arrastrar y soltar ───────────────────────────────────────────────────
    def _setup_dnd(self, sec_varias, sec_obras):
        """
        Registra todos los elementos de la ventana como destino de archivos.
        Los que estan dentro de una seccion fuerzan el tipo de esa seccion.
        """
        tipos = {str(sec_varias): 'varias', str(sec_obras): 'obras'}

        def recorrer(widget, tipo):
            tipo = tipos.get(str(widget), tipo)
            self._register_drop(widget, tipo, recursivo=False)
            for hijo in widget.winfo_children():
                recorrer(hijo, tipo)

        for hijo in self.winfo_children():
            recorrer(hijo, None)

    def _register_drop(self, widget, tipo, recursivo=True):
        """Hace que se puedan soltar archivos sobre el widget (y sus hijos)."""
        if not self._dnd:
            return
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind('<<Drop>>', lambda e, t=tipo: self._on_drop(e, t))
        except Exception:
            pass
        if recursivo:
            for hijo in widget.winfo_children():
                self._register_drop(hijo, tipo)

    def _on_drop(self, event, tipo):
        """
        Archivos soltados. Sobre una seccion van a esa seccion; sobre el resto
        de la ventana se clasifican por nombre (VARIAS -> Varias Obras).
        """
        paths   = [p for p in self.tk.splitlist(event.data)]
        excels  = [p for p in paths if p.lower().endswith(EXCEL_EXT)]
        ignored = len(paths) - len(excels)

        if tipo == 'varias':
            varias, obras = excels[:1], []
            ignored += len(excels[1:])
        elif tipo == 'obras':
            varias, obras = [], excels
        else:
            varias = [p for p in excels if es_reporte_varias(p)]
            obras  = [p for p in excels if not es_reporte_varias(p)]
            ignored += len(varias[1:])
            varias = varias[:1]

        if varias:
            self._add_varias(varias[0])
        if obras:
            self._add_obras(obras)
        if ignored:
            self.lbl_status.config(
                text=f"Se han ignorado {ignored} archivo(s): solo se admite Excel "
                     f"y un único reporte de Varias Obras.", fg="#a05a00")
        return event.action

    def _remove_obra(self, path):
        for i, (p, _, frame) in enumerate(self._obras):
            if p == path:
                frame.destroy()
                self._obras.pop(i)
                break
        for i, (_, _, frame) in enumerate(self._obras):
            frame.grid_configure(row=i)
        if not self._obras:
            self._lbl_empty_obras.grid()
        self.update_idletasks()
        self.geometry("")

    # ── Diálogo de configuración ─────────────────────────────────────────────
    def _open_settings(self):
        global CONFIG_HORAS_DIA_ALERTA, CONFIG_PORCENTAJE_MINIMO

        dlg = tk.Toplevel(self)
        dlg.title("Ajustes de umbrales")
        dlg.resizable(False, False)
        dlg.configure(bg="#f0f4f8")
        dlg.transient(self)
        dlg.grab_set()
        try:
            dlg.iconbitmap(resource_path("logo_programa_multi.ico"))
        except Exception:
            pass

        tk.Label(dlg, text="Configurar umbrales de alerta",
                 font=("Arial", 12, "bold"), bg="#1F4E79", fg="white",
                 pady=8, padx=16).grid(row=0, column=0, columnspan=2, sticky="ew")

        body = tk.Frame(dlg, bg="#f0f4f8", padx=18, pady=14)
        body.grid(row=1, column=0, columnspan=2, sticky="ew")

        # ── Horas por día ───────────────────────────────────────────────────
        tk.Label(body,
                 text="Horas en un día para resaltar la celda (en rojo/negrita):",
                 font=("Arial", 10), bg="#f0f4f8", anchor="w"
                 ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 4))

        var_horas = tk.IntVar(value=CONFIG_HORAS_DIA_ALERTA)
        tk.Spinbox(body, from_=1, to=12, textvariable=var_horas,
                   width=6, font=("Arial", 10),
                   relief="solid", bd=1
                   ).grid(row=1, column=0, sticky="w", padx=(0, 6))
        tk.Label(body, text="horas", font=("Arial", 9, "italic"),
                 bg="#f0f4f8", fg="#666"
                 ).grid(row=1, column=1, sticky="w")

        # ── Porcentaje ─────────────────────────────────────────────────────
        tk.Label(body,
                 text="Porcentaje mínimo de H. Extra sobre coste total para marcar fila:",
                 font=("Arial", 10), bg="#f0f4f8", anchor="w"
                 ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(14, 4))

        var_pct = tk.IntVar(value=int(round(CONFIG_PORCENTAJE_MINIMO * 100)))
        tk.Spinbox(body, from_=5, to=100, increment=5, textvariable=var_pct,
                   width=6, font=("Arial", 10),
                   relief="solid", bd=1
                   ).grid(row=3, column=0, sticky="w", padx=(0, 6))
        tk.Label(body, text="%", font=("Arial", 9, "italic"),
                 bg="#f0f4f8", fg="#666"
                 ).grid(row=3, column=1, sticky="w")

        # ── Botones ────────────────────────────────────────────────────────
        btns = tk.Frame(dlg, bg="#f0f4f8", pady=10, padx=18)
        btns.grid(row=2, column=0, columnspan=2, sticky="ew")

        def _save():
            global CONFIG_HORAS_DIA_ALERTA, CONFIG_PORCENTAJE_MINIMO
            try:
                h = int(var_horas.get())
                p = float(var_pct.get()) / 100.0
                if not (1 <= h <= 12):
                    raise ValueError("Horas fuera de rango (1-12)")
                if not (0.05 <= p <= 1.0):
                    raise ValueError("Porcentaje fuera de rango (5-100)")
            except Exception as e:
                messagebox.showerror("Valor no válido",
                                     f"Comprueba los valores introducidos.\n\n{e}",
                                     parent=dlg)
                return
            CONFIG_HORAS_DIA_ALERTA  = h
            CONFIG_PORCENTAJE_MINIMO = p
            guardar_configuracion()
            dlg.destroy()
            self.lbl_status.config(
                text=f"Ajustes guardados: {h}h/día · {int(p*100)}% del coste",
                fg="#1a6b1a")

        def _cancel():
            dlg.destroy()

        tk.Button(btns, text="Guardar", font=("Arial", 10, "bold"),
                  bg="#4CAF50", fg="white", activebackground="#388E3C",
                  relief="flat", cursor="hand2", padx=16, pady=6,
                  command=_save
                  ).pack(side="right", padx=(6, 0))
        tk.Button(btns, text="Cancelar", font=("Arial", 10),
                  bg="#e0e0e0", relief="flat", cursor="hand2",
                  padx=14, pady=6,
                  command=_cancel
                  ).pack(side="right")

        # Centrar la ventana sobre la principal
        dlg.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width()  - dlg.winfo_width())  // 2
        y = self.winfo_rooty() + (self.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry(f"+{x}+{y}")

    # ── Generación del informe ───────────────────────────────────────────────
    def _run(self):
        if not self._varias_path and not self._obras:
            messagebox.showwarning(
                "Faltan reportes",
                "Selecciona al menos un reporte antes de generar el informe.")
            return

        self.progress.start(10)
        self.lbl_status.config(text="Leyendo reportes…", fg="#444")
        self.update()

        try:
            sheets = []
            if self._varias_path:
                sheets.append(read_varias_obras(self._varias_path))
            for path, sheet_label, _ in self._obras:
                sheets.append(read_obra_xlsm(path, sheet_label))

            # Nombre por defecto con la quincena: Reporte_Horas_Extras_Atrasos_2Q_Septiembre_2026
            periodo = periodo_de(sheets)
            nombre  = "Reporte_Horas_Extras_Atrasos"
            if periodo:
                nombre += "_" + periodo[0].replace(" ", "_")
            self.progress.stop()
            output = filedialog.asksaveasfilename(
                title="Guardar informe de horas extra y atrasos como…",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")],
                initialfile=nombre + ".xlsx")
            if not output:
                self.lbl_status.config(text="", fg="#444")
                return

            self.progress.start(10)
            self.lbl_status.config(text="Generando informe…", fg="#444")
            self.update()
            build_excel(sheets, output)

            total_w  = sum(len(sh[1]) for sh in sheets)
            total_he = sum(x['TOTAL H. EXTRA'] for sh in sheets for x in sh[1])

            self.lbl_status.config(
                text=f"✔ Listo — {len(sheets)} hojas · {total_w} trabajadores · {total_he:g} h extra",
                fg="#1a6b1a")

            if messagebox.askyesno(
                "Informe generado",
                f"El informe se ha guardado en:\n{output}\n\n¿Quieres abrirlo ahora?"):
                import os, subprocess, platform
                if platform.system() == "Windows":
                    os.startfile(output)
                elif platform.system() == "Darwin":
                    subprocess.call(["open", output])
                else:
                    subprocess.call(["xdg-open", output])

        except Exception as e:
            import traceback
            messagebox.showerror(
                "Error",
                f"Ha ocurrido un error al procesar los reportes:\n\n{e}\n\n{traceback.format_exc()}")
            self.lbl_status.config(text="Error al procesar.", fg="red")
        finally:
            self.progress.stop()


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    cargar_configuracion()
    app = App()
    app.mainloop()
