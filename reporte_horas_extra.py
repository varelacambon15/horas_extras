"""
Generador de Reporte de Horas Extra
Solicita reportes de obra uno a uno y genera un Excel con una hoja por reporte.
Columnas: Trabajador, Ajustero, Sal. Dia, Horas por dia (una col por dia), Total H. Extra, Coste H. Extra
"""

import sys, re, json, os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas", "openpyxl"])
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter


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

def read_varias_obras(path):
    """
    Lee el reporte de Varias Obras.
    Devuelve: (nombre_hoja, workers, total_coste_obra, total_trabajadores)
    """
    df_raw = pd.read_excel(path, header=None)
    date_row   = df_raw.iloc[0]
    header_row = df_raw.iloc[1]

    # Detectar los bloques por dia: para cada fecha, las posiciones de
    # ADMIN, AJUSTE, OBRA 1, EXTRA 1, OBRA 2, EXTRA 2
    day_blocks = []
    last_date  = None
    i = 0
    while i < len(header_row):
        raw = str(date_row[i]) if pd.notna(date_row[i]) else ''
        m = re.search(r'(\d{2}-\d{2}-\d{4})', raw)
        if m:
            last_date = m.group(1)
        h = str(header_row[i]).strip() if pd.notna(header_row[i]) else ''
        if last_date and h == 'ADMIN':
            # Bloque tipico: ADMIN, AJUSTE, OBRA 1, EXTRA 1, OBRA 2, EXTRA 2
            day_blocks.append({
                'fecha':  last_date,
                'admin':  i,
                'ajuste': i + 1,
                'obra1':  i + 2,
                'extra1': i + 3,
                'obra2':  i + 4,
                'extra2': i + 5,
            })
            i += 6
        else:
            i += 1

    df = pd.read_excel(path, header=1)
    df = df.dropna(subset=['OPERARIOS']).copy()
    df = df[~df['OPERARIOS'].astype(str).str.strip().str.upper().isin(['TOTALES','TOTAL'])]

    sal_col  = 'SAL. DIA'
    adj_col  = 'AJUSTERO'
    tot_col  = 'TOTAL HORAS EXTRA'
    cost_col = 'SAL. HORA'

    df[sal_col]  = pd.to_numeric(df.get(sal_col,  0), errors='coerce').fillna(0)
    df[tot_col]  = pd.to_numeric(df.get(tot_col,  0), errors='coerce').fillna(0)
    df[cost_col] = pd.to_numeric(df.get(cost_col, 0), errors='coerce').fillna(0)

    def _es_codigo_obra(v):
        try:
            n = int(v)
            return 10000 <= n <= 99999
        except Exception:
            return False

    workers = []
    for row_i, (_, row) in enumerate(df.iterrows()):
        raw_row = df_raw.iloc[row_i + 2]

        # ── Recopilar horas por obra y horas por dia ────────────────────────
        day_hours   = {}   # fecha -> horas extra del dia
        obra_hours  = {}   # codigo_obra -> total de horas (regulares + extras)
        extras_obra = {}   # codigo_obra -> horas EXTRA del trabajador en esa obra

        # Obra "por defecto" del trabajador (columna OBRA en el head)
        obra_default = row.get('OBRA')
        obra_default_cod = str(int(obra_default)) if _es_codigo_obra(obra_default) else None

        for block in day_blocks:
            # Dias regulares (admin + ajuste) -> obra por defecto, 8h por dia
            admin_v  = pd.to_numeric(raw_row.iloc[block['admin']],  errors='coerce') or 0
            ajuste_v = pd.to_numeric(raw_row.iloc[block['ajuste']], errors='coerce') or 0
            dias_reg = float(admin_v) + float(ajuste_v)
            if dias_reg > 0 and obra_default_cod:
                obra_hours[obra_default_cod] = obra_hours.get(obra_default_cod, 0) + dias_reg * 8

            # Extras en OBRA 1 / OBRA 2
            for obra_col, extra_col in [(block['obra1'], block['extra1']),
                                        (block['obra2'], block['extra2'])]:
                extra_val = pd.to_numeric(raw_row.iloc[extra_col], errors='coerce')
                if not pd.notna(extra_val) or extra_val <= 0:
                    continue
                day_hours[block['fecha']] = day_hours.get(block['fecha'], 0) + float(extra_val)
                obra_val = raw_row.iloc[obra_col]
                if _es_codigo_obra(obra_val):
                    cod = str(int(obra_val))
                    obra_hours[cod]  = obra_hours.get(cod, 0)  + float(extra_val)
                    extras_obra[cod] = extras_obra.get(cod, 0) + float(extra_val)

        # Top 2 obras por horas (regulares + extras)
        top = sorted(obra_hours.items(), key=lambda x: -x[1])[:2]
        obra_str = '/'.join(cod for cod, _ in top) if top else ''

        total_e      = row[tot_col]
        sal_hora     = row[cost_col]
        coste_e      = total_e * sal_hora
        # Coste total = IMPORTE TOTAL (bruto, antes de prestamo y deducciones)
        importe_total = pd.to_numeric(row.get('IMPORTE TOTAL', 0), errors='coerce') or 0

        if total_e > 0:
            workers.append({
                'TRABAJADOR':     str(row['OPERARIOS']).strip(),
                'AJUSTERO':       str(row.get(adj_col, '')).strip().rstrip('*').strip(),
                'TECNICO':        str(row.get('TECNICO OBRA', '')).strip(),
                'SUPERVISOR':     '',
                'OFICIO':         str(row.get('OFICIO', '')).strip(),
                'OBRA':           obra_str,
                'SAL. DIA':       row[sal_col],
                'COSTE HORA':     sal_hora,
                'COSTE H. EXTRA': coste_e,
                'TOTAL H. EXTRA': total_e,
                'TOTAL A PAGAR':  importe_total,
                '_days':          day_hours,
                '_extras_obra':   extras_obra,
            })

    # Total de la obra: suma de IMPORTE TOTAL de todos los trabajadores
    total_obra         = pd.to_numeric(df.get('IMPORTE TOTAL', 0), errors='coerce').fillna(0).sum()
    total_trabajadores = len(df)
    return ('Varias Obras', workers, float(total_obra), total_trabajadores)


def read_obra_xlsm(path, sheet_name):
    """
    Devuelve: (nombre_hoja=codigo_obra, workers)
    El nombre de hoja es el codigo extraido del nombre de fichero o de los datos.
    """
    df = pd.read_excel(path, header=3, sheet_name='Reporte')
    df = df.dropna(subset=['OPERARIOS']).copy()
    df = df[~df['OPERARIOS'].astype(str).str.strip().str.upper().str.startswith('TOTAL')]

    extra_cols = [c for c in df.columns
                  if 'EXTRA' in str(c) and 'TOTAL' not in str(c) and 'COSTE' not in str(c)]

    # Extraer fecha de cada columna extra (dia y mes pueden venir como
    # 1 o 2 digitos: "08/07/2026" o "12/7/2026")
    date_per_col = {}
    for col in extra_cols:
        m = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', str(col))
        if m:
            d, mo, y = m.groups()
            date_per_col[col] = f"{int(d):02d}-{int(mo):02d}-{y}"

    for col in extra_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    sal_col  = 'SALARIO X DIA'
    adj_col  = 'AJUSTERO'
    tot_col  = 'TOTAL H. EXTRAS'
    cost_col = 'COSTE H. EXTRAS'

    df[sal_col]  = pd.to_numeric(df.get(sal_col,  0), errors='coerce').fillna(0)
    if tot_col in df.columns:
        df[tot_col] = pd.to_numeric(df[tot_col], errors='coerce').fillna(0)
    else:
        df[tot_col] = df[extra_cols].sum(axis=1)
    if cost_col in df.columns:
        df[cost_col] = pd.to_numeric(df[cost_col], errors='coerce').fillna(0)
    else:
        df[cost_col] = df[tot_col] * (df[sal_col] / 8)

    sup_col  = 'SUPERVISOR'
    has_sup  = sup_col in df.columns

    # Coste total de la obra (suma de TOTAL A PAGAR de TODOS los trabajadores,
    # tengan o no horas extra, para poder calcular el % de extras sobre el total)
    total_obra = pd.to_numeric(df.get('TOTAL A PAGAR', 0), errors='coerce').fillna(0).sum()

    has_oficio = 'OFICIO' in df.columns

    workers = []
    for _, row in df.iterrows():
        total_e = row[tot_col]
        if total_e <= 0:
            continue
        day_hours = {}
        for col, date_str in date_per_col.items():
            val = row[col]
            if val > 0:
                day_hours[date_str] = day_hours.get(date_str, 0) + float(val)
        sal_dia   = row[sal_col]
        coste_hora = sal_dia / 8 if sal_dia else 0
        workers.append({
            'TRABAJADOR':     str(row['OPERARIOS']).strip(),
            'AJUSTERO':       str(row.get(adj_col, '')).strip(),
            'TECNICO':        '',
            'SUPERVISOR':     str(row[sup_col]).strip() if has_sup else '',
            'OFICIO':         str(row['OFICIO']).strip() if has_oficio else '',
            'SAL. DIA':       sal_dia,
            'COSTE HORA':     coste_hora,
            'COSTE H. EXTRA': row[cost_col],
            'TOTAL H. EXTRA': total_e,
            'TOTAL A PAGAR':  pd.to_numeric(row.get('TOTAL A PAGAR', 0), errors='coerce') or 0,
            '_days':          day_hours,
        })

    total_trabajadores = len(df)
    return (sheet_name, workers, float(total_obra), total_trabajadores)


def detect_obra_code(path):
    """Intenta extraer el codigo de obra del nombre del fichero."""
    name = re.sub(r'[^0-9]', ' ', str(path))
    nums = [n for n in re.findall(r'\d{5}', name)]
    return nums[0] if nums else None


# ---------------------------------------------------------------------------
# Escritura del Excel
# ---------------------------------------------------------------------------

# Paleta azul — visualmente intuitiva, accesible para consulta rapida
HDR_BG   = "1F4E79"   # azul oscuro (cabecera)
HDR_FG   = "FFFFFF"
DAY_BG   = "BDD7EE"   # azul claro para cols de dias
TOT_BG   = "9DC3E6"   # azul medio para fila total
ALERT_FG = "C00000"   # rojo intenso para alertas (texto)

def _brd_outer():
    """Borde exterior completo."""
    s = Side(style='thin', color='999999')
    return Border(left=s, right=s, top=s, bottom=s)

def _brd_row():
    """Solo linea inferior (separador de filas)."""
    b = Side(style='thin', color='CCCCCC')
    n = Side(style=None)
    return Border(bottom=b, left=n, right=n, top=n)

def _brd_none():
    n = Side(style=None)
    return Border(left=n, right=n, top=n, bottom=n)

def _hdr_font():  return Font(name='Calibri', bold=True, color=HDR_FG, size=10)
def _hdr_fill():  return PatternFill('solid', start_color=HDR_BG)
def _day_fill():  return PatternFill('solid', start_color=DAY_BG)
def _tot_fill():  return PatternFill('solid', start_color=TOT_BG)
def _cen():       return Alignment(horizontal='center', vertical='center')
def _left():      return Alignment(horizontal='left',   vertical='center')


# El umbral de alerta ahora se calcula dinamicamente por hoja en write_sheet,
# en base a la media de horas extra del propio reporte (ver alli).

def write_sheet(wb, sheet_name, workers):
    ws = wb.create_sheet(title=sheet_name[:31])
    money = '#,##0.00'
    dec1  = '0.0'
    pct   = '0.0%'

    if not workers:
        ws['A1'] = 'Sin trabajadores con horas extra'
        return

    all_days  = sorted(set(d for w in workers for d in w['_days']))
    has_sup   = any(w.get('SUPERVISOR') for w in workers)
    has_tec   = any(w.get('TECNICO')    for w in workers)
    has_ofi   = any(w.get('OFICIO')     for w in workers)
    has_obra  = any(w.get('OBRA')       for w in workers)

    # Umbral dinamico de horas extra (basado en media del propio reporte):
    #   - al menos 2.3x la media Y al menos 25h por encima de la media
    horas_lista = [w['TOTAL H. EXTRA'] for w in workers]
    media       = sum(horas_lista) / len(horas_lista) if horas_lista else 0
    umbral_h    = max(media * 2.3, media + 25)

    # Umbral dinamico de % H. Extra sobre coste total:
    #   - al menos 1.5x el % medio Y al menos 10 puntos porcentuales por encima
    #   - y nunca por debajo del 25% absoluto (mas restrictivo)
    porcentajes = []
    for w in workers:
        if w.get('TOTAL A PAGAR', 0) > 0:
            porcentajes.append(w['COSTE H. EXTRA'] / w['TOTAL A PAGAR'])
    media_pct  = sum(porcentajes) / len(porcentajes) if porcentajes else 0
    umbral_pct = max(media_pct * 1.5, media_pct + 0.10, CONFIG_PORCENTAJE_MINIMO)

    fixed = ['TRABAJADOR', 'AJUSTERO']
    if has_tec:
        fixed.append('TECNICO')
    if has_sup:
        fixed.append('SUPERVISOR')
    if has_ofi:
        fixed.append('OFICIO')
    if has_obra:
        fixed.append('OBRA')
    fixed += ['SAL. DIA', 'COSTE HORA']
    # Cola: costes primero, luego horas, luego %  (horas extras justo antes del %)
    tail  = ['COSTE TOTAL', 'COSTE H. EXTRA', 'TOTAL H. EXTRA', '% H. EXTRA']
    all_cols = fixed + all_days + tail
    n_cols   = len(all_cols)

    # ── Fila 1: cabecera ─────────────────────────────────────────────────────
    LABELS = {
        'TRABAJADOR':     'Trabajador',
        'AJUSTERO':       'Ajustero',
        'TECNICO':        'Técnico Obra',
        'SUPERVISOR':     'Supervisor',
        'OFICIO':         'Oficio',
        'OBRA':           'Obra',
        'SAL. DIA':       'Sal. Dia',
        'COSTE HORA':     'Coste/Hora',
        'TOTAL H. EXTRA': 'Total H. Extra',
        'COSTE TOTAL':    'Coste Total',
        'COSTE H. EXTRA': 'Coste H. Extra',
        '% H. EXTRA':     '% H. Extra',
    }

    for ci, col in enumerate(all_cols, 1):
        label = LABELS.get(col, col)
        c = ws.cell(1, ci, label)
        c.font      = _hdr_font()
        if col in all_days:
            c.fill = _day_fill()
            c.font = Font(name='Calibri', bold=True, color=HDR_BG, size=9)
        else:
            c.fill = _hdr_fill()
        c.alignment = _cen()
        c.border    = _brd_outer()
    ws.row_dimensions[1].height = 20

    # ── Filas de datos ───────────────────────────────────────────────────────
    for ri, w in enumerate(workers, 2):
        total_e        = w['TOTAL H. EXTRA']
        coste_e        = w['COSTE H. EXTRA']
        coste_total    = w.get('TOTAL A PAGAR', 0)
        pct_extra      = (coste_e / coste_total) if coste_total > 0 else 0

        # Doble criterio: muchas horas extra O alto % sobre coste total
        es_alerta_h    = total_e > umbral_h
        es_alerta_pct  = pct_extra > umbral_pct and pct_extra > 0
        es_alerta      = es_alerta_h or es_alerta_pct

        for ci, col in enumerate(all_cols, 1):
            if col in ('TRABAJADOR', 'AJUSTERO', 'TECNICO', 'SUPERVISOR', 'OFICIO', 'OBRA'):
                val = w.get(col, '')
            elif col == 'SAL. DIA':
                val = w['SAL. DIA']
            elif col == 'COSTE HORA':
                val = w['COSTE HORA']
            elif col == 'TOTAL H. EXTRA':
                val = total_e
            elif col == 'COSTE TOTAL':
                val = coste_total
            elif col == 'COSTE H. EXTRA':
                val = coste_e
            elif col == '% H. EXTRA':
                val = pct_extra
            else:
                val = w['_days'].get(col, '')

            bold  = col in ('TOTAL H. EXTRA', '% H. EXTRA') and es_alerta
            color = ALERT_FG if es_alerta else '000000'
            # Resaltado adicional: celdas individuales de dia con mas de 6h
            if col in all_days:
                try:
                    if float(val) > CONFIG_HORAS_DIA_ALERTA:
                        color = ALERT_FG
                        bold  = True
                except (TypeError, ValueError):
                    pass
            c = ws.cell(ri, ci, val if val != 0.0 else '')
            c.font      = Font(name='Calibri', size=10, bold=bold, color=color)
            c.border    = _brd_row()
            c.alignment = _left() if col in ('TRABAJADOR', 'AJUSTERO', 'TECNICO', 'SUPERVISOR', 'OFICIO') else _cen()
            if col in ('SAL. DIA', 'COSTE HORA', 'COSTE TOTAL', 'COSTE H. EXTRA'):
                c.number_format = money
            elif col == 'TOTAL H. EXTRA' or col in all_days:
                c.number_format = dec1
            elif col == '% H. EXTRA':
                c.number_format = pct

    # ── Fila totales ─────────────────────────────────────────────────────────
    last_data = 1 + len(workers)
    tr        = last_data + 1
    for ci, col in enumerate(all_cols, 1):
        c = ws.cell(tr, ci)
        c.fill      = _tot_fill()
        c.font      = Font(name='Calibri', bold=True, size=10)
        c.border    = _brd_outer()
        c.alignment = _left() if ci == 1 else _cen()
        cl = get_column_letter(ci)
        if ci == 1:
            c.value = 'Total'
        elif col in all_days or col == 'TOTAL H. EXTRA':
            c.value         = f'=SUM({cl}2:{cl}{last_data})'
            c.number_format = dec1
        elif col in ('COSTE TOTAL', 'COSTE H. EXTRA'):
            c.value         = f'=SUM({cl}2:{cl}{last_data})'
            c.number_format = money
        elif col == '% H. EXTRA':
            # % global de la obra
            col_e = get_column_letter(all_cols.index('COSTE H. EXTRA') + 1)
            col_t = get_column_letter(all_cols.index('COSTE TOTAL')    + 1)
            c.value         = f'=IFERROR({col_e}{tr}/{col_t}{tr},0)'
            c.number_format = pct

    # ── Anchos ───────────────────────────────────────────────────────────────
    col_widths = {
        'TRABAJADOR': 30, 'AJUSTERO': 26, 'TECNICO': 18, 'SUPERVISOR': 20, 'OFICIO': 14, 'OBRA': 14,
        'SAL. DIA': 10, 'COSTE HORA': 11,
        'TOTAL H. EXTRA': 13, 'COSTE TOTAL': 14, 'COSTE H. EXTRA': 14, '% H. EXTRA': 12,
    }
    for ci, col in enumerate(all_cols, 1):
        w = col_widths.get(col, 10)
        ws.column_dimensions[get_column_letter(ci)].width = w

    # ── Filtros ──────────────────────────────────────────────────────────────
    # Excluye la fila de totales del filtro
    last_col_letter = get_column_letter(n_cols)
    ws.auto_filter.ref = f"A1:{last_col_letter}{last_data}"

    ws.freeze_panes = 'A2'


def build_excel(sheets, output_path):
    """sheets es lista de (sheet_name, workers, coste_total_obra, total_trabajadores)."""
    wb = Workbook()
    wb.remove(wb.active)

    # ── Hoja resumen ──────────────────────────────────────────────────────────
    # Columnas: Obra | Trabajadores | Trab. c/HE | Coste Total | Coste H. Extra | Total H. Extra | %
    ws_r  = wb.create_sheet('Resumen', 0)
    money = '#,##0.00'
    pct   = '0.0%'

    headers = ['Obra', 'Trabajadores', 'Trab. con HE', 'Coste Total',
               'Coste H. Extra', 'Total H. Extra', '% s/Total']
    for ci, label in enumerate(headers, 1):
        c = ws_r.cell(1, ci, label)
        c.font      = _hdr_font()
        c.fill      = _hdr_fill()
        c.alignment = _cen()
        c.border    = _brd_outer()
    ws_r.row_dimensions[1].height = 18

    for ri, (sname, workers, coste_total_obra, total_trab) in enumerate(sheets, 2):
        total_he    = sum(w['TOTAL H. EXTRA'] for w in workers)
        total_cost  = sum(w['COSTE H. EXTRA'] for w in workers)
        porcentaje  = (total_cost / coste_total_obra) if coste_total_obra > 0 else 0

        valores = [sname, total_trab, len(workers), coste_total_obra,
                   total_cost, total_he, porcentaje]
        for ci, val in enumerate(valores, 1):
            c = ws_r.cell(ri, ci, val)
            c.font      = Font(name='Calibri', size=10)
            c.border    = _brd_row()
            c.alignment = _left() if ci == 1 else _cen()
            if ci == 4: c.number_format = money
            if ci == 5: c.number_format = money
            if ci == 6: c.number_format = '0.0'
            if ci == 7: c.number_format = pct

    tr = 1 + len(sheets) + 1
    for ci in range(1, 8):
        c = ws_r.cell(tr, ci)
        c.fill      = _tot_fill()
        c.font      = Font(name='Calibri', bold=True, size=10)
        c.border    = _brd_outer()
        c.alignment = _left() if ci == 1 else _cen()
        cl = get_column_letter(ci)
        if   ci == 1: c.value = 'Total'
        elif ci == 2: c.value = f'=SUM({cl}2:{cl}{tr-1})'
        elif ci == 3: c.value = f'=SUM({cl}2:{cl}{tr-1})'
        elif ci == 4: c.value = f'=SUM({cl}2:{cl}{tr-1})'; c.number_format = money
        elif ci == 5: c.value = f'=SUM({cl}2:{cl}{tr-1})'; c.number_format = money
        elif ci == 6: c.value = f'=SUM({cl}2:{cl}{tr-1})'; c.number_format = '0.0'
        elif ci == 7:
            # % global = total coste extras / total coste obras
            c.value = f'=IFERROR(E{tr}/D{tr},0)'
            c.number_format = pct

    ws_r.column_dimensions['A'].width = 22
    ws_r.column_dimensions['B'].width = 13
    ws_r.column_dimensions['C'].width = 13
    ws_r.column_dimensions['D'].width = 16
    ws_r.column_dimensions['E'].width = 16
    ws_r.column_dimensions['F'].width = 15
    ws_r.column_dimensions['G'].width = 12
    ws_r.auto_filter.ref = f"A1:G{tr-1}"
    ws_r.freeze_panes = 'A2'

    for sname, workers, _, _ in sheets:
        write_sheet(wb, sname, workers)

    wb.save(output_path)


# ---------------------------------------------------------------------------
# Interfaz grafica
# ---------------------------------------------------------------------------




def resource_path(filename):
    """Devuelve la ruta al recurso, compatible con PyInstaller (--onefile)."""
    import os
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, filename)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
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
            text="Aún no se ha añadido el reporte de Varias Obras.",
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
            text="Aún no se ha añadido ningún reporte de obra.",
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
        return row

    # ── Acciones: Varias Obras ───────────────────────────────────────────────
    def _add_varias(self):
        p = filedialog.askopenfilename(
            title="Selecciona el reporte de Varias Obras",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
        if not p:
            return
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
    def _add_obras(self):
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

        output = filedialog.asksaveasfilename(
            title="Guardar informe de horas extra como…",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="Reporte_Horas_Extra.xlsx")
        if not output:
            return

        self.progress.start(10)
        self.lbl_status.config(text="Procesando…", fg="#444")
        self.update()

        try:
            sheets = []
            if self._varias_path:
                name, workers, total_obra, total_trab = read_varias_obras(self._varias_path)
                sheets.append((name, workers, total_obra, total_trab))
            for path, sheet_label, _ in self._obras:
                name, workers, total_obra, total_trab = read_obra_xlsm(path, sheet_label)
                sheets.append((name, workers, total_obra, total_trab))

            build_excel(sheets, output)

            total_w  = sum(len(w) for _, w, _, _ in sheets)
            total_he = sum(x['TOTAL H. EXTRA'] for _, w, _, _ in sheets for x in w)

            self.lbl_status.config(
                text=f"✔ Listo — {len(sheets)} hojas · {total_w} trabajadores · {total_he:.1f} h extra",
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
