# -*- coding: utf-8 -*-

#Bibliotecas necessarias
#---------------------------------------------------------------------------------------------------------------------
import pandas as pd
import numpy as np
import warnings
from datetime import datetime as dt
import os
import io
import math
import textwrap
from matplotlib import pyplot as plt
from matplotlib.ticker import FuncFormatter
from matplotlib import colors as mcolors
from matplotlib.patches import Rectangle
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
from decimal import Decimal, ROUND_CEILING, ROUND_DOWN, ROUND_FLOOR, ROUND_HALF_DOWN, ROUND_HALF_EVEN, ROUND_HALF_UP, ROUND_UP, ROUND_05UP
from pptx import Presentation 
from pptx.util import Inches, Cm, Pt
from pptx.dml.color import RGBColor
from typing import NamedTuple, Optional
from pathlib import Path
from collections import OrderedDict

COLOR_BLUE = '\033[94m'
COLOR_YELLOW = '\033[93m'
COLOR_GREEN = '\033[92m'
COLOR_RED = '\033[91m'
COLOR_RESET = '\033[0m'

COLOR_QUESTION = '\033[38;5;37m'

HEADER_COLOR_PRIMARY = '#286B72'
HEADER_COLOR_SECONDARY = '#3EBBC7'
HEADER_FONT_COLOR = '#FFFFFF'
HEADER_TOTAL_FILL = '#E8F0FF'
HEADER_FIRST_COL_FILL = '#F5F5F5'
TABLE_TEXT_PRIMARY = '#1F1F1F'
TABLE_POSITIVE_COLOR = '#27AE60'
TABLE_NEGATIVE_COLOR = '#C0392B'
TABLE_GRID_COLOR = '#1A1A1A'
APORTE_NEGATIVE_COLOR = "#F3948B"
APORTE_NEUTRAL_COLOR = "#FFCE7E"
VOLUME_BAR_START = '#1452B8'
VOLUME_BAR_END = '#8CB8FF'
VOLUME_BAR_ALPHA = 0.85
VOLUME_BAR_BASE_ALPHA = 0.18
LINE_COLOR_CLIENT = '#2C3E50'
LINE_COLOR_NUMERATOR = "#D4AC0D"
BAR_COLOR_VARIATION_CLIENT = '#7F8C8D'
BAR_COLOR_VARIATION_NUMERATOR = '#F1C40F'
BAR_EDGE_COLOR = 'black'
ANNOTATION_BOX_FACE = '#F2F2F2'
ANNOTATION_BOX_EDGE = 'black'

TREND_COLOR_PALETTE = {
    'trend_01': '#1F77B4',
    'trend_02': '#FF7F0E',
    'trend_03': '#2CA02C',
    'trend_04': '#D62728',
    'trend_05': '#9467BD',
    'trend_06': '#8C564B',
    'trend_07': '#E377C2',
    'trend_08': '#7F7F7F',
    'trend_09': '#BCBD22',
    'trend_10': '#17BECF',
    'trend_11': '#AEC7E8',
    'trend_12': '#FFBB78',
    'trend_13': '#98DF8A',
    'trend_14': '#FF9896',
    'trend_15': '#C5B0D5',
    'trend_16': '#C49C94',
    'trend_17': '#F7B6D2',
    'trend_18': '#C7C7C7',
    'trend_19': '#DBDB8D',
    'trend_20': '#9EDAE5',
    'trend_21': '#393B79',
    'trend_22': '#5254A3',
    'trend_23': '#6B6ECF',
    'trend_24': '#9C9EDE',
    'trend_25': '#637939',
    'trend_26': '#8CA252',
    'trend_27': '#B5CF6B',
    'trend_28': '#CEDB9C',
    'trend_29': '#8C6D31',
    'trend_30': '#BD9E39',
    'trend_31': '#E7BA52',
    'trend_32': '#E7CB94',
    'trend_33': '#843C39',
    'trend_34': '#AD494A',
    'trend_35': '#D6616B',
    'trend_36': '#E7969C',
    'trend_37': '#7B4173',
    'trend_38': '#A55194',
    'trend_39': '#CE6DBD',
    'trend_40': '#DE9ED6',
    'trend_41': '#3182BD',
    'trend_42': '#6BAED6',
    'trend_43': '#9ECAE1',
    'trend_44': '#C6DBEF',
    'trend_45': '#E6550D',
    'trend_46': '#FD8D3C',
    'trend_47': '#FDAE6B',
    'trend_48': '#FDD0A2',
    'trend_49': '#31A354',
    'trend_50': '#74C476',
    'trend_51': '#A1D99B',
    'trend_52': '#C7E9C0',
    'trend_53': '#756BB1',
    'trend_54': '#9E9AC8',
    'trend_55': '#BCBDDC',
    'trend_56': '#DADAEB',
    'trend_57': '#636363',
    'trend_58': '#969696',
    'trend_59': '#BDBDBD',
    'trend_60': '#D9D9D9',
    'trend_61': '#393E46',
    'trend_62': '#00ADB5',
    'trend_63': '#FF5722',
    'trend_64': '#795548',
    'trend_65': '#607D8B',
    'trend_66': '#8BC34A',
    'trend_67': '#CDDC39',
    'trend_68': '#FFC107',
    'trend_69': '#FF4081',
    'trend_70': '#3F51B5'
}
TREND_COLOR_SEQUENCE = list(TREND_COLOR_PALETTE.values())
TABLE_WRAP_WIDTH = 14
DISPLAY_TREND_REFERENCE_TEXT = False

TREND_SCALE_RULES = [
    {
        "threshold": 1_000_000_000_000,
        "divisor": 1_000_000_000_000,
        "suffix": "T",
        "reference": {
            'E': 'Valores expresados en billones',
            'P': 'Valores expressos em trilhões',
            'default': 'Values expressed in trillions'
        }
    },
    {
        "threshold": 1_000_000_000,
        "divisor": 1_000_000_000,
        "suffix": "B",
        "reference": {
            'E': 'Valores expresados en miles de millones',
            'P': 'Valores expressos em bilhões',
            'default': 'Values expressed in billions'
        }
    },
    {
        "threshold": 1_000_000,
        "divisor": 1_000_000,
        "suffix": "M",
        "reference": {
            'E': 'Valores expresados en millones',
            'P': 'Valores expressos em milhões',
            'default': 'Values expressed in millions'
        }
    },
    {
        "threshold": 1_000,
        "divisor": 1_000,
        "suffix": "K",
        "reference": {
            'E': 'Valores expresados en miles',
            'P': 'Valores expressos em milhares',
            'default': 'Values expressed in thousands'
        }
    },
    {
        "threshold": 0,
        "divisor": 1,
        "suffix": "",
        "reference": {
            'E': '',
            'P': '',
            'default': ''
        }
    },
]


MONTH_NAMES = {
    'P': ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
    'E': ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
}


def colorize(text: str, color: str = COLOR_BLUE) -> str:
    return f"{color}{text}{COLOR_RESET}"


def print_colored(text: str, color: str = COLOR_BLUE) -> None:
    print(colorize(text, color))


def select_trend_scale_rule(max_value: float) -> dict:
    for rule in TREND_SCALE_RULES:
        if max_value >= rule["threshold"]:
            return rule
    return TREND_SCALE_RULES[-1]


def format_value_with_suffix(value: float, divisor: float, suffix: str) -> str:
    if not np.isfinite(value):
        return ''
    scaled = value / divisor if divisor else value
    abs_scaled = abs(scaled)
    if abs_scaled >= 100:
        formatted = f"{scaled:.0f}"
    elif abs_scaled >= 10:
        formatted = f"{scaled:.1f}"
    else:
        formatted = f"{scaled:.2f}"
    if '.' in formatted:
        formatted = formatted.rstrip('0').rstrip('.')
    return f"{formatted}{suffix}"


def available_width(presentation, left=Inches(0), right=Inches(0)):
    return presentation.slide_width - int(left) - int(right)


def constrain_picture_width(picture, max_width):
    """Clamp picture width to max_width while preserving aspect ratio."""
    if max_width is None:
        return
    max_width = int(max_width)
    if max_width <= 0:
        return
    current_width = int(picture.width)
    if current_width <= max_width:
        return
    current_height = int(picture.height)
    scale = current_width / max_width
    picture.width = max_width
    picture.height = max(int(round(current_height / scale)), 1)


EMU_PER_INCH = 914400
DEFAULT_LINE_CHART_RATIO = 3
TABLE_TARGET_HEIGHT_CM = 4.0
TABLE_SIDE_MARGIN_CM = 1.2
TABLE_PAIR_GAP_CM = 0.5
TABLE_HEADER_FONT_SIZE = 10
TABLE_WRAP_WIDTH = 14

LINE_CHART_LEFT_MARGIN = 0.045
LINE_CHART_RIGHT_MARGIN = 0.99
LINE_CHART_BOTTOM_MARGIN = 0.1
LINE_CHART_TOP_MARGIN = 0.9
LINE_CHART_SINGLE_X_MARGIN = 0.16

LINE_CHART_MULTI_LEFT_MARGIN = 0.08
LINE_CHART_MULTI_RIGHT_MARGIN = 0.98
LINE_CHART_MULTI_BOTTOM_MARGIN = 0.12
LINE_CHART_MULTI_TOP_MARGIN = 0.85
LINE_CHART_MULTI_X_MARGIN = 0.16


def emu_to_inches(value: int) -> float:
    return float(value) / EMU_PER_INCH


def wrap_table_text(value, max_width: int = TABLE_WRAP_WIDTH):
    """
    Return the given value with soft line breaks so long strings wrap nicely inside table cells.
    Keeps non-string values intact.
    """
    if not isinstance(value, str):
        return value
    normalized = " ".join(value.split())
    if not normalized:
        return ""
    if len(normalized) <= max_width or " " not in normalized:
        return normalized
    return textwrap.fill(normalized, width=max_width, break_long_words=False)


pd.set_option('future.no_silent_downcasting', True)
pd.set_option('mode.chained_assignment', None)
warnings.filterwarnings('ignore')

agora = dt.now()

#Funcao que prepara os dados para criação do gráfico MAT
def df_mat(df,p):
    
    v1 = pd.DataFrame([(df.iloc[i-12-p:i-p,1].sum()/df.iloc[i-24-p:i-12-p,1].sum()) - 1 if i >= 24 else np.nan for i in range(12, len(df)+1)],columns=['Var Sell-in'])
    v2 = pd.DataFrame([(df.iloc[i-12:i,2].sum()/df.iloc[i-24:i-12,2].sum()) - 1 if i >= 24 else np.nan for i in range(12, len(df)+1)],columns=['Var Sell-out'])

    ac1, ac2 = [pd.DataFrame(df[col].rolling(window=12).sum().dropna()) for col in df.columns[1:]]
    v1.index,v2.index=v1.index+11,v2.index+11
    v1 ,v2 = v1.dropna(), v2.dropna()

    mat =pd.DataFrame(df.iloc[:,0].copy())
    mat= mat.join(ac1.join(ac2)[-len(v1):].join([v1]).join([v2]))
    mat= mat[mat.notna().all(axis=1).idxmax():]

    return mat

#Grafico de MAT 
def graf_mat (mat,c_fig,p):

    #Cria figura e área plotável   
        altura = 7  # Aumentamos aún más la altura
        # Cerrar cualquier figura existente con el mismo número para que figsize sea aplicado
        try:
            plt.close(c_fig)
        except Exception:
            pass
        fig = plt.figure(num=c_fig, figsize=(15, altura))
        ax = fig.add_subplot(1, 1, 1)
        # Ajustamos los márgenes para dar más espacio arriba y abajo
        plt.subplots_adjust(bottom=0.2, top=0.9)
        
        #Cria o eixo do tempo e os acumulados juntamente com as variações
        ran=mat.iloc[:,0].copy()
        ran=[x.strftime('%m-%y') for x in ran]
        ac1 = mat.iloc[:,1].copy()
        ac2 = mat.iloc[:,2].copy()
        v1= mat.iloc[:,3].copy()
        v2= mat.iloc[:,4].copy()

        #Plota os acumulados em linhas    
        l1=ax.plot(ran,ac1, color=LINE_COLOR_CLIENT, linewidth=4.2, label='Acumulado Cliente')
        l2=ax.plot(ran,ac2, color=LINE_COLOR_NUMERATOR, linewidth=4.2, label='Acumulado Numerator') 

        #Plota as variações em barras
        ax2= ax.twinx()
        b1 = ax2.bar(np.arange(len(ran)) - 0.3, v1.values, 0.3, color=BAR_COLOR_VARIATION_CLIENT, edgecolor=BAR_EDGE_COLOR, label='Var % Cliente Pipeline: '+ str(p) +' '+labels[(lang,'Var MAT')])
        b2 = ax2.bar(np.arange(len(ran)) + 0.3, v2.values, 0.3, color=BAR_COLOR_VARIATION_NUMERATOR, edgecolor=BAR_EDGE_COLOR, label='Var % Numerator '+labels[(lang,'Var MAT')])

        ax.set_xticks(np.arange(len(ran)))
        ax.set_xticklabels(ran, rotation=30)
        ax2.tick_params(left=False, labelleft=False, top=False, labeltop=False,
            right=False, labelright=False, bottom=False, labelbottom=False)

        for v in [v1,v2]:
            for x,y in zip(np.arange(len(v1))+0.2,v):
                label = "{:.1f}%".format(y)
                bbox_props_white=dict(facecolor=ANNOTATION_BOX_FACE,edgecolor=ANNOTATION_BOX_EDGE)
                plt.annotate(f"{y*100:.1f}%", (x, y), textcoords="offset points", xytext=(0, 10), ha='center', color='red' if y < 0 else 'green', size=9, bbox=bbox_props_white)

        
        #Cria a legenda 
        lns = l1 + l2 + [b1, b2]
        labs = [l.get_label() for l in l1 + l2] + [b1.get_label(), b2.get_label()]
        ax.legend(lns, labs, loc='upper center', bbox_to_anchor=(0.5, -0.15), borderaxespad=0.05, frameon=True, prop={'size': 12}, ncol=2)
        # Ajustamos el layout para evitar que se corten elementos
        plt.tight_layout()
        img_stream = io.BytesIO()


        #Cria o titulo
        ax.set_title(labels[(lang,'MAT')]+' | ' + w[2:], size=18, pad=20)
        
        # Ajustamos el layout antes de guardar la imagen
        plt.tight_layout(rect=[0, 0.05, 1, 0.95])  # Ajusta el espacio dejando margen para título y leyenda

        #Salva o grafico
        # Guardar respetando el tamaño definido (no usar bbox_inches='tight')
        fig.savefig(img_stream, format='png', transparent=True)

        #Insere o grafico
        img_stream.seek(0) 

        # Cerrar figura para liberar recursos y evitar reutilización
        try:
            plt.close(fig)
        except Exception:
            pass

        return img_stream


#Grafico de Lineas
def line_graf(
    df,
    p,
    title,
    c_fig,
    ven,
    width_emu=None,
    height_emu=None,
    multi_chart=None,
    share_lookup=None,
    y_axis_percent=False,
    top_value_annotations=0,
    color_collector=None,
    color_overrides=None,
    linestyle_overrides=None,
):
    """
    Renderiza la gráfica de tendencias y devuelve un buffer PNG listo para insertar en la slide.
    Cuando se generan múltiples gráficos en la misma diapositiva (multi_chart=True), se aplican
    márgenes más amplios para evitar recortes visuales.
    """
    if width_emu is not None:
        try:
            planned_width = emu_to_inches(width_emu)
        except Exception:
            planned_width = None
    else:
        planned_width = None

    if multi_chart is None:
        detected_multi = planned_width is not None and planned_width < 6.0
    else:
        detected_multi = bool(multi_chart)

    if detected_multi:
        chart_top_margin = LINE_CHART_MULTI_TOP_MARGIN
        chart_bottom_margin = LINE_CHART_MULTI_BOTTOM_MARGIN
        chart_left_margin = LINE_CHART_MULTI_LEFT_MARGIN
        chart_right_margin = LINE_CHART_MULTI_RIGHT_MARGIN
        chart_x_margin = LINE_CHART_MULTI_X_MARGIN
    else:
        chart_top_margin = LINE_CHART_TOP_MARGIN
        chart_bottom_margin = LINE_CHART_BOTTOM_MARGIN
        chart_left_margin = LINE_CHART_LEFT_MARGIN
        chart_right_margin = LINE_CHART_RIGHT_MARGIN
        chart_x_margin = LINE_CHART_SINGLE_X_MARGIN

    # Ajusta o tamanho da figura para combinar com o espaco reservado no slide
    width_inches = emu_to_inches(width_emu) if width_emu is not None else None
    height_inches = emu_to_inches(height_emu) if height_emu is not None else None

    if width_inches is None and height_inches is None:
        # Ancho dinámico según cantidad de datos (en pulgadas)
        n_points = len(df)
        ancho = max(15, n_points * 0.5)  # 0.5 puede ajustarse para más/menos espacio (pulgadas)
        # Altura fija de 10 cm -> convertir a pulgadas
        altura = emu_to_inches(Cm(10))
        figsize = (ancho, altura)
    else:
        if width_inches is None:
            width_inches = height_inches * DEFAULT_LINE_CHART_RATIO
        elif height_inches is None:
            height_inches = width_inches / DEFAULT_LINE_CHART_RATIO
        figsize = (width_inches, height_inches)

    fig = plt.figure(num=c_fig, figsize=figsize)
    ax = fig.add_subplot(1, 1, 1)
    # Escalar elementos (fuentes, linewidth) según la altura para que el contenido se adapte
    # Se toma como referencia la altura original de 5 pulgadas usada previamente
    ref_height = 5.0
    actual_height = figsize[1]
    scale = actual_height / ref_height if ref_height else 1.0
    base_linewidth = 2.0 * max(0.6, scale)
    title_base_size = max(10, int(18 * scale))
    legend_base_size = max(8, int(10 * scale))
    xtick_size = max(8, int(10 * scale))
    marker_size = max(3.5, 4.5 * scale)

    aux = df.copy()

    if aux.shape[1] > 1:
        kept_columns = [aux.columns[0]]
        for col in aux.columns[1:]:
            series = pd.to_numeric(aux[col], errors='coerce')
            if series.replace(0, np.nan).dropna().empty:
                continue
            kept_columns.append(col)
        if len(kept_columns) == 1:
            aux = aux.iloc[:, :1]
        else:
            aux = aux.loc[:, kept_columns]

    if aux.empty or aux.shape[1] <= 1:
        img_stream = io.BytesIO()
        fig.savefig(img_stream, format="png", transparent=True)
        img_stream.seek(0)
        try:
            plt.close(fig)
        except Exception:
            pass
        return img_stream

    ran = [x.strftime('%m-%y') for x in aux.iloc[:, 0]]
    data_len = len(ran)
    if data_len == 0:
        img_stream = io.BytesIO()
        fig.savefig(img_stream, format="png", transparent=True)
        img_stream.seek(0)
        try:
            plt.close(fig)
        except Exception:
            pass
        return img_stream

    start_idx = max(0, min(p, data_len - 1))
    x_positions = np.arange(data_len)

    colunas = list(aux.columns[1:])
    palette_values = TREND_COLOR_SEQUENCE or [mcolors.to_hex(c) for c in plt.get_cmap('tab20').colors]
    color_mapping = {}
    if color_overrides:
        for name, hex_color in color_overrides.items():
            color_mapping[name] = hex_color
    color_index = 0

    if ven > 1:
        comp = colunas[:len(colunas) // 2]
        vent = colunas[len(colunas) // 2:]
        pair_count = min(len(comp), len(vent))
        for idx in range(pair_count):
            color = palette_values[color_index % len(palette_values)]
            color_index += 1
            color_mapping[comp[idx]] = color
            color_mapping[vent[idx]] = color
        for base in comp[pair_count:]:
            if base not in color_mapping:
                color_mapping[base] = palette_values[color_index % len(palette_values)]
                color_index += 1
        for base in vent[pair_count:]:
            if base not in color_mapping:
                color_mapping[base] = palette_values[color_index % len(palette_values)]
                color_index += 1
    else:
        for base in colunas:
            if base not in color_mapping:
                color_mapping[base] = palette_values[color_index % len(palette_values)]
                color_index += 1

    lns = []
    legend_labels = []
    numeric_series_list = []
    plotted_columns = []
    series_points = {}
    for col in colunas:
        if ven > 1:
            estilo = '-' if '.v' in col.lower() else '--'
        else:
            estilo = '-'

        cor = color_mapping.get(col, palette_values[0])
        if color_overrides and col in color_overrides:
            cor = color_overrides[col]
            color_mapping[col] = cor

        numeric_series = pd.to_numeric(aux[col], errors='coerce')
        numeric_series_list.append(numeric_series)
        y = numeric_series.values
        x_slice = x_positions[start_idx:]
        y_slice = y[start_idx:]
        valid_points = [(x_idx, y_val) for x_idx, y_val in zip(x_slice, y_slice) if pd.notna(y_val)]
        if not valid_points:
            continue
        series_points[col] = valid_points

        suffix_map = {".c": "Compras", ".v": "Ventas", "_c": "Compras", "_v": "Ventas", "-c": "Compras", "-v": "Ventas"}
        lower_col = col.lower()
        legend_label = col
        for suffix, suffix_text in suffix_map.items():
            if lower_col.endswith(suffix):
                legend_label = f"{col[:-len(suffix)].strip()} ({suffix_text})"
                break

        if linestyle_overrides and col in linestyle_overrides:
            estilo = linestyle_overrides[col]

        line, = ax.plot(
            x_slice,
            y_slice,
            color=cor,
            linewidth=base_linewidth,
            linestyle=estilo,
            label=col,
            marker='o',
            markersize=marker_size,
            markerfacecolor='white',
            markeredgewidth=max(0.7, base_linewidth / 2),
            markeredgecolor=cor,
        )
        lns.append(line)
        legend_labels.append(legend_label)
        plotted_columns.append(col)


    if data_len and start_idx < data_len:
        ax.set_xticks(x_positions[start_idx:])
        ax.set_xticklabels(ran[start_idx:], rotation=30, fontsize=xtick_size)
        ax.set_xlim(x_positions[start_idx], x_positions[-1])
    else:
        ax.set_xticks([])
        ax.set_xticklabels([])

    y_tick_size = max(8, int(10 * scale))
    ax.tick_params(axis='y', labelsize=y_tick_size)

    axis_percent = bool(y_axis_percent)
    max_abs_value = 0.0
    for series in numeric_series_list:
        numeric_values = series.to_numpy()
        if numeric_values.size == 0:
            continue
        finite_values = numeric_values[np.isfinite(numeric_values)]
        if finite_values.size == 0:
            continue
        series_max = float(np.max(np.abs(finite_values)))
        if series_max > max_abs_value:
            max_abs_value = series_max

    if axis_percent:
        def _trend_tick_formatter(value, _):
            return f"{value:.1f}%"

        ax.yaxis.set_major_formatter(FuncFormatter(_trend_tick_formatter))
    else:
        scale_rule = select_trend_scale_rule(max_abs_value)
        divisor = scale_rule["divisor"] or 1
        suffix = scale_rule["suffix"]

        def _trend_tick_formatter(value, _):
            return format_value_with_suffix(value, divisor, suffix)

        ax.yaxis.set_major_formatter(FuncFormatter(_trend_tick_formatter))

        if DISPLAY_TREND_REFERENCE_TEXT and suffix:
            reference_template = scale_rule.get("reference", {})
            reference_text = reference_template.get(lang, reference_template.get('default', '')).strip()
            if reference_text:
                reference_font_size = max(8, int(9 * scale))
                ax.text(
                    0.02,
                    0.95,
                    reference_text,
                    transform=ax.transAxes,
                    ha='left',
                    va='top',
                    fontsize=reference_font_size,
                    color='#4D4D4D',
                    alpha=0.85,
                )

    if axis_percent:
        def _format_axis_value(value):
            return f"{value:.1f}%"
    else:
        def _format_axis_value(value):
            return format_value_with_suffix(value, divisor, suffix)

    ax.set_ylim(bottom=0)
    ax.margins(x=chart_x_margin, y=0.08)

    legend_gap_base = 0.025 if detected_multi else 0.02
    legend_bottom_margin = chart_bottom_margin
    legend_columns = 1
    legend_rows = 1
    legend_height_fraction = 0.0
    margin_buffer = 0.01
    if lns:
        max_columns = 3 if detected_multi else 4
        legend_columns = max(1, min(len(legend_labels), max_columns))
        legend_rows = max(1, math.ceil(len(legend_labels) / legend_columns))
        legend_font_points = legend_base_size
        legend_line_height_points = legend_font_points * 1.35
        legend_height_inches = (legend_line_height_points / 72.0) * legend_rows
        figure_height_inches = fig.get_size_inches()[1]
        legend_height_fraction = legend_height_inches / figure_height_inches if figure_height_inches else 0.0
        legend_bottom_margin = max(
            chart_bottom_margin,
            legend_height_fraction + legend_gap_base
        ) + margin_buffer

    plt.title(title, size=title_base_size, pad=10)
    bottom_margin = min(0.9, max(0.05, legend_bottom_margin))
    if bottom_margin >= chart_top_margin:
        bottom_margin = max(0.05, chart_top_margin - 0.05)
    fig.subplots_adjust(
        top=chart_top_margin,
        bottom=bottom_margin,
        left=chart_left_margin,
        right=chart_right_margin,
    )

    if lns:
        axes_height = ax.get_position().height
        axes_height = axes_height if axes_height > 0 else 1e-3
        legend_gap_fraction = legend_gap_base
        legend_offset_axes = legend_gap_fraction / axes_height
        legend_offset_axes = max(0.0, min(0.35, legend_offset_axes))
        legend = ax.legend(
            lns,
            legend_labels,
            loc='upper left',
            bbox_to_anchor=(0.0, -legend_offset_axes, 1.0, 0.0),
            borderaxespad=0.0,
            frameon=True,
            prop={'size': legend_base_size},
            ncol=legend_columns,
            mode='expand',
        )
        frame = legend.get_frame()
        frame.set_facecolor('white')
        frame.set_edgecolor('#D3D3D3')
        frame.set_alpha(0.85)

        try:
            if fig.canvas is None:
                FigureCanvas(fig)
        except Exception:
            pass

        adjustment_iterations = 0
        while adjustment_iterations < 4:
            adjustment_iterations += 1
            try:
                fig.canvas.draw()
                renderer = fig.canvas.get_renderer()
            except Exception:
                break

            figure_height_px = fig.get_size_inches()[1] * fig.dpi
            if not figure_height_px:
                break

            legend_bbox = legend.get_window_extent(renderer)
            tick_bboxes = [
                label.get_window_extent(renderer)
                for label in ax.get_xticklabels()
                if label.get_text()
            ]

            tick_height_fraction = 0.0
            min_tick_bottom = None
            if tick_bboxes:
                tick_height_fraction = max(bbox.height for bbox in tick_bboxes) / figure_height_px
                min_tick_bottom = min(bbox.y0 for bbox in tick_bboxes)

            clearance_px = max(4.0, 4.0 * scale)
            clearance_fraction = clearance_px / figure_height_px
            desired_gap_fraction = legend_gap_base + tick_height_fraction + clearance_fraction

            if min_tick_bottom is not None:
                overlap_px = legend_bbox.y1 - (min_tick_bottom - clearance_px)
                if overlap_px > 0:
                    desired_gap_fraction += overlap_px / figure_height_px

            required_bottom_margin = max(chart_bottom_margin, legend_height_fraction + desired_gap_fraction) + margin_buffer
            if legend_bbox.y0 < 0:
                required_bottom_margin = max(
                    required_bottom_margin,
                    bottom_margin + (-legend_bbox.y0 + clearance_px) / figure_height_px
                )

            required_bottom_margin = min(0.9, required_bottom_margin)
            if required_bottom_margin >= chart_top_margin:
                required_bottom_margin = max(0.05, chart_top_margin - 0.05)

            axes_height = ax.get_position().height
            axes_height = axes_height if axes_height > 0 else 1e-3
            new_offset_axes = desired_gap_fraction / axes_height
            new_offset_axes = max(0.0, min(0.45, new_offset_axes))

            no_layout_change = (
                abs(required_bottom_margin - bottom_margin) < 1e-4 and
                abs(new_offset_axes - legend_offset_axes) < 1e-4
            )

            bottom_margin = required_bottom_margin
            fig.subplots_adjust(
                top=chart_top_margin,
                bottom=bottom_margin,
                left=chart_left_margin,
                right=chart_right_margin,
            )

            axes_height = ax.get_position().height
            axes_height = axes_height if axes_height > 0 else 1e-3
            legend_offset_axes = max(0.0, min(0.45, desired_gap_fraction / axes_height))
            legend.set_bbox_to_anchor((0.0, -legend_offset_axes, 1.0, 0.0))

            if no_layout_change:
                break

    if plotted_columns and share_lookup:
        legend_label_map = {col: label for col, label in zip(plotted_columns, legend_labels)}
        annotation_candidates = []
        share_lookup_lower = {str(key).lower(): value for key, value in share_lookup.items()} if share_lookup else {}
        for col, line in zip(plotted_columns, lns):
            points = series_points.get(col)
            if not points:
                continue
            last_x, last_y = points[-1]
            share_value = share_lookup.get(col)
            if share_value is None:
                share_value = share_lookup_lower.get(str(col).lower())
            annotation_candidates.append({
                "column": col,
                "line": line,
                "legend": legend_label_map.get(col, col),
                "share": share_value,
                "last_x": last_x,
                "last_y": last_y,
                "color": line.get_color(),
            })

        total_series = len(annotation_candidates)
        if total_series:
            if total_series < 10:
                selected_series = annotation_candidates
            else:
                with_share = [info for info in annotation_candidates if info["share"] is not None]
                without_share = [info for info in annotation_candidates if info["share"] is None]
                with_share.sort(key=lambda info: info["share"], reverse=True)
                selected_series = with_share[:5]
                if len(selected_series) < 5:
                    remaining = [info for info in annotation_candidates if info not in selected_series]
                    selected_series.extend(remaining[: max(0, 5 - len(selected_series))])

            if selected_series:
                x_limits = ax.get_xlim()
                y_limits = ax.get_ylim()
                y_range = y_limits[1] - y_limits[0] if y_limits[1] > y_limits[0] else 1.0
                min_gap = y_range * 0.04
                min_margin = y_range * 0.02
                x_range = x_limits[1] - x_limits[0] if x_limits[1] > x_limits[0] else 1.0
                x_offset = max(0.8, x_range * 0.06)

                arranged_series = sorted(selected_series, key=lambda info: info["last_y"], reverse=True)
                placed_positions = []
                y_max = y_limits[1] - min_margin
                y_min = y_limits[0] + min_margin
                for info in arranged_series:
                    y_pos = max(min(info["last_y"], y_max), y_min)
                    for prev_y in placed_positions:
                        if y_pos > prev_y - min_gap:
                            y_pos = prev_y - min_gap
                    y_pos = max(y_pos, y_min)
                    placed_positions.append(y_pos)
                    info["label_y"] = y_pos
                    info["label_x"] = info["last_x"] + x_offset

                max_label_x = max(info["label_x"] for info in arranged_series)
                if max_label_x > x_limits[1]:
                    ax.set_xlim(x_limits[0], max_label_x + x_offset * 1.15)

                size_factor = 0.7 if detected_multi else 0.8
                min_size = 6 if detected_multi else 7
                label_font_size = max(min_size, int(legend_base_size * size_factor))
                bbox_props = dict(facecolor='white', edgecolor='none', alpha=0.85, boxstyle='round,pad=0.2')
                for info in arranged_series:
                    text_value = info["legend"]
                    ax.text(
                        info["label_x"],
                        info["label_y"],
                        text_value,
                        ha='left',
                        va='center',
                        color=info["color"],
                        fontsize=label_font_size,
                        fontweight='bold',
                        bbox=bbox_props,
                        clip_on=False,
                    )

    if top_value_annotations and plotted_columns and lns:
        point_candidates = []
        line_lookup = {col: line for col, line in zip(plotted_columns, lns)}
        for col in plotted_columns:
            if str(col).strip().lower() == 'total':
                continue
            line = line_lookup.get(col)
            if line is None:
                continue
            for x_idx, y_val in series_points.get(col, []):
                if not np.isfinite(y_val):
                    continue
                point_candidates.append({
                    "column": col,
                    "line": line,
                    "x": x_idx,
                    "y": float(y_val),
                })
        if point_candidates:
            point_candidates.sort(key=lambda item: item["y"], reverse=True)
            y_limits = ax.get_ylim()
            y_range = y_limits[1] - y_limits[0] if y_limits[1] > y_limits[0] else 1.0
            label_offset = max(y_range * 0.03, 0.5 if axis_percent else y_range * 0.03)
            label_font_size = max(8, int(10 * scale))
            annotated = 0
            seen_points = set()
            for point in point_candidates:
                point_key = (point["x"], round(point["y"], 6))
                if point_key in seen_points:
                    continue
                seen_points.add(point_key)
                label_y = point["y"] + label_offset
                if label_y > y_limits[1]:
                    label_y = min(y_limits[1], point["y"] + label_offset * 0.6)
                text_value = _format_axis_value(point["y"])
                ax.annotate(
                    text_value,
                    xy=(point["x"], point["y"]),
                    xytext=(point["x"], label_y),
                    textcoords='data',
                    ha='center',
                    va='bottom',
                    fontsize=label_font_size,
                    color=point["line"].get_color(),
                    fontweight='bold',
                    bbox=dict(
                        boxstyle='round,pad=0.2',
                        facecolor='white',
                        edgecolor=point["line"].get_color(),
                        linewidth=0.6,
                        alpha=0.85
                    ),
                    arrowprops=dict(
                        arrowstyle='-',
                        color=point["line"].get_color(),
                        linewidth=max(0.6, base_linewidth * 0.5)
                    ),
                    clip_on=False,
                )
                annotated += 1
                if annotated >= top_value_annotations:
                    break

    if isinstance(color_collector, dict):
        color_collector.clear()
        for col, line in zip(plotted_columns, lns):
            color_collector[col] = line.get_color()

    img_stream = io.BytesIO()
    fig.savefig(img_stream, format='png', transparent=True)
    img_stream.seek(0)

    try:
        plt.close(fig)
    except Exception:
        pass

    return img_stream


#Normaliza etiquetas de periodo eliminando prefijos genéricos
def normalize_period_label(label) -> str:
    if label is None:
        return ""
    text = str(label).strip()
    if not text:
        return text
    parts = text.split(None, 1)
    if parts and parts[0].lower().rstrip('.') in {'vol', 'volume', 'volumen'} and len(parts) > 1:
        return parts[1].strip()
    return text


#Grafico de barras apiladas 100% para share por periodo
def stacked_share_chart(period_label, share_values, color_mapping, c_fig, title=None):
    """
    Genera un gráfico 100% apilado a partir de la participación de un periodo puntual.
    """
    if not isinstance(share_values, (dict, OrderedDict)):
        share_values = {}

    palette_values = TREND_COLOR_SEQUENCE or [mcolors.to_hex(c) for c in plt.get_cmap('tab20').colors]

    normalized_items = []
    total = 0.0
    for key, value in share_values.items():
        try:
            numeric_value = float(value)
        except (TypeError, ValueError):
            try:
                numeric_value = float(str(value).replace('%', '').replace(',', '.')) / 100.0
            except Exception:
                numeric_value = 0.0
        if not np.isfinite(numeric_value) or numeric_value < 0:
            numeric_value = 0.0
        total += numeric_value
        normalized_items.append([key, numeric_value])

    if total <= 0:
        total = 1.0

    for item in normalized_items:
        item[1] = item[1] / total

    normalized_items = [item for item in normalized_items if item[1] > 0]

    segments = len(normalized_items)
    base_height = 6.2
    incremental_height = 0.6
    figure_height = max(base_height, 4.5 + segments * incremental_height)
    figure_height = min(7.0, figure_height)
    figure_width = 4.6
    fig, ax = plt.subplots(num=c_fig, figsize=(figure_width, figure_height))

    bottom = 0.0
    bar_width = 0.65
    label_infos = []
    for idx, (key, value) in enumerate(normalized_items):
        color = color_mapping.get(key)
        if color is None:
            color = palette_values[idx % len(palette_values)]
            color_mapping[key] = color
        ax.bar(0, value, width=bar_width, bottom=bottom, color=color, edgecolor='white')
        center_y = bottom + value / 2 if value > 0 else bottom
        share_label = f"{value * 100:.1f}%"
        label_infos.append(
            {
                "key": key,
                "text": f"{key} ({share_label})",
                "color": color,
                "center_y": center_y,
            }
        )
        bottom += value

    if label_infos:
        label_infos.sort(key=lambda info: info["center_y"])
        min_gap = max(0.035, min(0.12, 0.95 / max(segments, 1)))
        min_y = 0.04
        max_y = 0.96
        adjusted_positions = []
        for info in label_infos:
            proposed = np.clip(info["center_y"], min_y, max_y)
            if adjusted_positions and proposed - adjusted_positions[-1] < min_gap:
                proposed = min(max_y, adjusted_positions[-1] + min_gap)
            adjusted_positions.append(proposed)
        for idx in range(len(adjusted_positions) - 2, -1, -1):
            if adjusted_positions[idx + 1] - adjusted_positions[idx] < min_gap:
                adjusted_positions[idx] = max(min_y, adjusted_positions[idx + 1] - min_gap)
        for info, position in zip(label_infos, adjusted_positions):
            info["label_y"] = position

        max_text_len = max(len(info["text"]) for info in label_infos)
        label_offset = 1.1 + max(0, (max_text_len - 22) * 0.014)
        label_offset = min(label_offset, 1.8)
        label_x = bar_width / 2 + label_offset
    else:
        label_x = bar_width / 2 + 1.15

    connector_x = bar_width / 2 + 0.1
    label_font_size = 9 if segments <= 10 else (8 if segments <= 16 else 7)
    if label_infos:
        for info in label_infos:
            ax.plot(
                [connector_x, label_x - 0.1],
                [info["center_y"], info["label_y"]],
                color=info["color"],
                linewidth=1.2,
                alpha=0.65,
            )
            ax.text(
                label_x,
                info["label_y"],
                info["text"],
                ha='left',
                va='center',
                color=info["color"],
                fontsize=label_font_size,
                fontweight='bold'
            )

    ax.set_ylim(0, 1)
    axis_right = label_x + 1.0 if label_infos else (bar_width / 2 + 1.4)
    ax.set_xlim(-0.75, axis_right)
    ax.set_xticks([])
    ax.set_ylabel('')
    ax.yaxis.set_major_formatter(FuncFormatter(lambda x, _: f"{x * 100:.0f}%"))
    ax.set_title(title or period_label, fontsize=15, pad=24, fontweight='bold')
    ax.set_yticks(np.linspace(0, 1, 5))
    ax.yaxis.tick_left()
    ax.grid(axis='y', linestyle='--', alpha=0.25)

    for spine in ax.spines.values():
        spine.set_visible(False)

    fig.subplots_adjust(left=0.05, right=0.92, top=0.995, bottom=0.02)
    img_stream = io.BytesIO()
    fig.savefig(img_stream, format='png', transparent=True, bbox_inches=None)
    img_stream.seek(0)

    try:
        plt.close(fig)
    except Exception:
        pass

    return img_stream, fig.get_size_inches()

#Função que cria a tabela de aporte
def aporte(df,p,lang,tipo):

        aux = df.copy()

        aux['Total'] = aux.iloc[:len(df),1:].sum(axis=1)

        apo=pd.DataFrame(columns=[tipo] + aux.columns[1:].tolist())
        #Vol ultimo MAT
        apo.loc[len(apo)] = [str('Vol' )+" "+aux.loc[len(aux)-1-12-p,labels[(lang,'Data')]].strftime('%b-%y') ] + [aux.iloc[len(aux)-24-p:len(aux)-12-p, col].sum() / aux.iloc[len(aux)-24-p:len(aux)-12-p,aux.shape[1]-1].sum() for col in range(1,len(aux.columns) )]
        #Vol MAT atual
        apo.loc[len(apo)] = [str('Vol' )+" "+aux.loc[len(aux)-1-p,labels[(lang,'Data')]].strftime('%b-%y') ] + [aux.iloc[len(aux)-12-p:len(aux)-p, col].sum() / aux.iloc[len(aux)-12-p:len(aux)-p,aux.shape[1]-1].values.sum() for col in range(1,len(aux.columns))]
        #Var
        apo.loc[len(apo)] = ["Var %"] +  [ aux.iloc[len(aux)-12-p:len(aux)-p, col].sum() / aux.iloc[len(aux)-24-p:len(aux)-12-p, col].sum()-1 for col in range(1,len(aux.columns))]
        #Aporte
        apo.loc[len(apo)] = ["Aporte"] + [apo.iloc[-1, col] * apo.iloc[0, col] for col in range(1,len(aux.columns))]

        # Remove categories with invalid metrics (NaN/Inf) before formatting
        def _is_invalid_metric(value):
            if pd.isna(value):
                return True
            try:
                numeric = float(value)
            except (TypeError, ValueError):
                return False
            return not np.isfinite(numeric)

        invalid_columns = []
        for col_idx in range(1, len(apo.columns)):
            col_name = apo.columns[col_idx]
            var_value = apo.iloc[2, col_idx]
            aporte_value = apo.iloc[3, col_idx]
            if _is_invalid_metric(var_value) or _is_invalid_metric(aporte_value):
                invalid_columns.append(col_name)

        if invalid_columns:
            apo.drop(columns=invalid_columns, inplace=True)

        apo.attrs["removed_columns"] = list(dict.fromkeys(invalid_columns))

        share_period_values = []
        max_share_rows = min(2, len(apo.index))
        for row_idx in range(max_share_rows):
            period_label = normalize_period_label(apo.iloc[row_idx, 0])
            period_shares = OrderedDict()
            for col_idx in range(1, len(apo.columns)):
                column_name = str(apo.columns[col_idx])
                normalized_name = column_name.strip().lower()
                if not normalized_name or normalized_name == 'total':
                    continue
                value = apo.iloc[row_idx, col_idx]
                numeric_value = None
                try:
                    numeric_value = float(value)
                except (TypeError, ValueError):
                    try:
                        numeric_value = float(str(value).replace('%', '').replace(',', '.')) / 100.0
                    except Exception:
                        numeric_value = None
                if numeric_value is None or not np.isfinite(numeric_value) or numeric_value < 0:
                    numeric_value = 0.0
                period_shares[column_name] = numeric_value
            share_period_values.append((period_label, period_shares))

        apo.attrs["share_period_values"] = share_period_values

        share_mat_values = {}
        if len(share_period_values) > 1:
            _, current_period_shares = share_period_values[1]
            share_mat_values = {
                column_name: value
                for column_name, value in current_period_shares.items()
                if np.isfinite(value)
            }
        apo.attrs["share_mat_values"] = share_mat_values

        #Formatação do volume
        apo.iloc[:2, 1:] = apo.iloc[:2, 1:].applymap(lambda x: f"{round(x * 100, 1)}%")
        #Formatação da variação e aporte
        apo.iloc[2:, 1:] = apo.iloc[2:, 1:].applymap(lambda x: f"{round(x * 100, 2)}%")

        return apo

#Função que cria o gráfico tabela de aporte

def graf_apo(apo, c_fig, column_color_mapping=None):
    fig, ax = plt.subplots(num=c_fig)
    row_height = 0.45
    col_width = 1.0

    n_rows, n_cols = apo.shape
    fig_width = col_width * n_cols
    fig_height = row_height * (n_rows + 0.2)
    fig.set_size_inches(fig_width, fig_height)
    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
    ax.axis('off')
    ax.set_facecolor('white')

    display_data = apo.copy()
    if display_data.shape[1] > 0:
        display_data.iloc[:, 0] = display_data.iloc[:, 0].map(wrap_table_text)
    display_col_labels = [wrap_table_text(str(label)) for label in apo.columns]

    table = ax.table(
        cellText=display_data.values,
        colLabels=display_col_labels,
        loc='center',
        cellLoc='center'
    )

    for i, _ in enumerate(apo.columns):
        table.auto_set_column_width(i)

    table.scale(1, 1.2)
    table.auto_set_font_size(False)
    table.set_fontsize(10)

    header_main = HEADER_COLOR_PRIMARY
    header_secondary = HEADER_COLOR_SECONDARY
    total_fill = HEADER_TOTAL_FILL
    first_col_fill = HEADER_FIRST_COL_FILL
    positive_color = TABLE_POSITIVE_COLOR
    negative_color = TABLE_NEGATIVE_COLOR
    bar_padding_ratio = 0.08
    bar_height_ratio = 0.55

    column_color_mapping = column_color_mapping or {}
    exact_color_mapping = {
        str(key).strip(): value
        for key, value in column_color_mapping.items()
        if key is not None and value
    }
    normalized_color_mapping = {
        key.lower(): value
        for key, value in exact_color_mapping.items()
        if key
    }

    suffix_variants = ('.c', '.v', '_c', '_v', '-c', '-v')

    def resolve_column_color(column_name):
        if column_name is None:
            return None
        key = str(column_name).strip()
        if not key:
            return None
        color_value = exact_color_mapping.get(key)
        if color_value:
            return color_value
        lower_key = key.lower()
        color_value = normalized_color_mapping.get(lower_key)
        if color_value:
            return color_value
        for suffix in suffix_variants:
            if lower_key.endswith(suffix):
                base = key[:-len(suffix)].strip()
                if not base:
                    continue
                color_value = exact_color_mapping.get(base)
                if color_value:
                    return color_value
                color_value = normalized_color_mapping.get(base.lower())
                if color_value:
                    return color_value
        return None

    total_column_name = next((col for col in apo.columns if str(col).strip().lower() == 'total'), None)

    data_columns = [
        idx
        for idx in range(1, n_cols)
        if total_column_name is None or apo.columns[idx] != total_column_name
    ]
    start_rgb = np.array(mcolors.to_rgb(VOLUME_BAR_START))
    end_rgb = np.array(mcolors.to_rgb(VOLUME_BAR_END))

    def share_colors(count: int):
        if count <= 0:
            return []
        if count == 1:
            return [mcolors.to_hex(np.clip(start_rgb, 0, 1))]
        colors = []
        for i in range(count):
            ratio = i / (count - 1)
            rgb = np.clip(start_rgb + (end_rgb - start_rgb) * ratio, 0, 1)
            colors.append(mcolors.to_hex(rgb))
        return colors

    fallback_column_colors = share_colors(len(data_columns))
    column_colors = []
    for rel_idx, col_idx in enumerate(data_columns):
        resolved = resolve_column_color(apo.columns[col_idx])
        if resolved is None:
            resolved = fallback_column_colors[rel_idx] if rel_idx < len(fallback_column_colors) else VOLUME_BAR_END
        try:
            column_colors.append(mcolors.to_hex(resolved))
        except (ValueError, TypeError):
            column_colors.append(fallback_column_colors[rel_idx] if rel_idx < len(fallback_column_colors) else VOLUME_BAR_END)
    white_rgb = np.array([1.0, 1.0, 1.0])
    volume_row_indexes = [idx for idx, label in enumerate(apo.iloc[:, 0]) if str(label).lower().startswith('vol')]
    aporte_row_index = next((idx for idx, label in enumerate(apo.iloc[:, 0]) if str(label).strip().lower() == 'aporte'), None)

    def to_percentage(value):
        try:
            text_value = str(value).strip().replace('%', '').replace(',', '.')
            if not text_value:
                return None
            return max(0.0, float(text_value) / 100.0)
        except Exception:
            return None

    def to_signed_percentage(value):
        try:
            text_value = str(value).strip().replace('%', '').replace(',', '.')
            if not text_value:
                return None
            return float(text_value) / 100.0
        except Exception:
            return None

    for (row, col), cell in table.get_celld().items():
        cell.set_edgecolor(TABLE_GRID_COLOR)
        cell.set_linewidth(1)
        cell.get_text().set_zorder(3)

        if row == 0:
            text = cell.get_text()
            text.set_weight('bold')
            text.set_fontsize(TABLE_HEADER_FONT_SIZE)
            header_text_color = HEADER_FONT_COLOR
            if col == 0 or (total_column_name is not None and apo.columns[col] == total_column_name):
                cell.set_facecolor(header_main)
            else:
                header_color = resolve_column_color(apo.columns[col])
                if header_color is None:
                    cell.set_facecolor(header_secondary)
                else:
                    try:
                        cell.set_facecolor(header_color)
                        rgb = mcolors.to_rgb(header_color)
                        luminance = 0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2]
                        header_text_color = '#FFFFFF' if luminance < 0.6 else TABLE_TEXT_PRIMARY
                    except (ValueError, TypeError):
                        cell.set_facecolor(header_secondary)
            text.set_color(header_text_color)
        else:
            label = str(apo.iloc[row - 1, 0])
            text = cell.get_text()

            if col == 0:
                cell.set_facecolor(first_col_fill)
                text.set_weight('bold')
                text.set_color(TABLE_TEXT_PRIMARY)
            else:
                if total_column_name is not None and apo.columns[col] == total_column_name:
                    cell.set_facecolor(total_fill)
                else:
                    cell.set_facecolor('white')

                if label.lower().startswith('vol'):
                    text.set_color(TABLE_TEXT_PRIMARY)
                else:
                    value = apo.iloc[row - 1, col]
                    try:
                        numeric = float(str(value).replace('%', '').replace(',', '.'))
                        text.set_color(positive_color if numeric >= 0 else negative_color)
                    except Exception:
                        pass

    if aporte_row_index is not None:
        aporte_values = []
        for col in data_columns:
            numeric = to_signed_percentage(apo.iloc[aporte_row_index, col])
            if numeric is not None:
                aporte_values.append((col, numeric))
        if aporte_values:
            aporte_numbers = [value for _, value in aporte_values]
            min_val = min(aporte_numbers)
            max_val = max(aporte_numbers)
            cmap = None
            norm = None
            if not math.isclose(min_val, max_val, rel_tol=1e-9, abs_tol=1e-9):
                if min_val < 0 and max_val > 0:
                    cmap = mcolors.LinearSegmentedColormap.from_list(
                        'aporte_diverging',
                        [APORTE_NEGATIVE_COLOR, APORTE_NEUTRAL_COLOR, TABLE_POSITIVE_COLOR]
                    )
                    norm = mcolors.TwoSlopeNorm(vmin=min_val, vcenter=0, vmax=max_val)
                elif max_val <= 0:
                    cmap = mcolors.LinearSegmentedColormap.from_list(
                        'aporte_negative',
                        [APORTE_NEGATIVE_COLOR, APORTE_NEUTRAL_COLOR]
                    )
                    norm = mcolors.Normalize(vmin=min_val, vmax=0)
                else:
                    cmap = mcolors.LinearSegmentedColormap.from_list(
                        'aporte_positive',
                        [APORTE_NEUTRAL_COLOR, TABLE_POSITIVE_COLOR]
                    )
                    norm = mcolors.Normalize(vmin=0, vmax=max_val)
            for col, numeric in aporte_values:
                if cmap is not None and norm is not None:
                    rgba = cmap(norm(numeric))
                    rgb = tuple(rgba[:3])
                else:
                    if math.isclose(numeric, 0.0, abs_tol=1e-9):
                        rgb = mcolors.to_rgb(APORTE_NEUTRAL_COLOR)
                    elif numeric > 0:
                        rgb = mcolors.to_rgb(TABLE_POSITIVE_COLOR)
                    else:
                        rgb = mcolors.to_rgb(APORTE_NEGATIVE_COLOR)
                cell = table[(aporte_row_index + 1, col)]
                cell.set_facecolor(rgb)
                text = cell.get_text()
                luminance = 0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2]
                text.set_color(TABLE_TEXT_PRIMARY if luminance > 0.5 else '#FFFFFF')

    for df_row_idx in volume_row_indexes:
        table_row = df_row_idx + 1
        row_percentages = {}
        max_percent = 0.0
        for col in data_columns:
            percent = to_percentage(apo.iloc[df_row_idx, col])
            row_percentages[col] = percent
            if percent is not None and percent > max_percent:
                max_percent = percent
        for rel_idx, col in enumerate(data_columns):
            percent = row_percentages.get(col)
            if percent is None:
                continue
            cell = table[(table_row, col)]
            bar_color = column_colors[rel_idx] if rel_idx < len(column_colors) else VOLUME_BAR_END
            base_rgb = np.array(mcolors.to_rgb(bar_color))
            intensity = min(percent / max_percent, 1.0) if max_percent > 0 else 0.0
            blended_rgb = white_rgb * (1.0 - intensity) + base_rgb * intensity
            cell.set_facecolor(blended_rgb)
            cell.get_text().set_color(TABLE_TEXT_PRIMARY)

    fig.canvas.draw()
    renderer = fig.canvas.get_renderer()
    table_bbox = table.get_tightbbox(renderer).transformed(fig.dpi_scale_trans.inverted())
    target_height_in = TABLE_TARGET_HEIGHT_CM / 2.54
    if table_bbox.height > 0:
        scale = target_height_in / table_bbox.height
        if abs(scale - 1.0) > 1e-3:
            current_width, current_height = fig.get_size_inches()
            fig.set_size_inches(current_width * scale, current_height * scale)
            fig.canvas.draw()
            renderer = fig.canvas.get_renderer()
            table_bbox = table.get_tightbbox(renderer).transformed(fig.dpi_scale_trans.inverted())
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=400, bbox_inches=table_bbox, pad_inches=0)
    buf.seek(0)
    return buf

def simplify_name_segment(value: str, max_len: int) -> str:
    """
    Reduce a string to a compact, filesystem-friendly segment.
    Keeps alphanumeric characters, replaces others with dashes, and trims length.
    """
    if value is None:
        return 'NA'
    value_str = str(value).strip()
    if not value_str:
        return 'NA'
    cleaned = ''.join(ch if ch.isalnum() else '-' for ch in value_str)
    cleaned = cleaned.strip('-')
    if not cleaned:
        cleaned = 'NA'
    return cleaned[:max_len]

def select_excel_file(base_dir: Path) -> str:
    excel_files = sorted([p for p in base_dir.glob('*.xlsx') if not p.name.startswith('~$')])

    def metadata_for(file_path: Path) -> str:
        try:
            parts = file_path.stem.split('_')
            country_code = int(parts[0]) if parts else None
            category_code = parts[1] if len(parts) > 1 else None
            pais_name = pais.loc[pais['cod'] == country_code, 'pais'].iloc[0] if country_code is not None else 'Pais N/D'
            cesta = categ.loc[categ['cod'] == category_code, 'cest'].iloc[0] if category_code else 'Cesta N/D'
            categoria = categ.loc[categ['cod'] == category_code, 'cat'].iloc[0] if category_code else 'Categoria N/D'
            return f"{pais_name} | {cesta} | {categoria}"
        except Exception:
            return 'Metadata no disponible'

    if not excel_files:
        raise FileNotFoundError(colorize(f'No se encontraron archivos .xlsx en {base_dir}', COLOR_RED))

    if len(excel_files) == 1:
        meta = metadata_for(excel_files[0])
        print_colored(f"Archivo encontrado: {excel_files[0].name} {colorize('[' + meta + ']', COLOR_YELLOW)}", COLOR_BLUE)
        return excel_files[0].name

    print_colored('Archivos Excel disponibles:', COLOR_QUESTION)
    for idx, archivo in enumerate(excel_files, 1):
        meta = metadata_for(archivo)
        print(f"{colorize(f'  {idx}. {archivo.name}', COLOR_BLUE)} {colorize('[' + meta + ']', COLOR_YELLOW)}")

    prompt = colorize(f"Seleccione el numero del archivo (1-{len(excel_files)}). Enter para {excel_files[0].name}: ", COLOR_QUESTION)
    while True:
        choice = input(prompt).strip()
        if not choice:
            return excel_files[0].name
        if choice.isdigit():
            idx = int(choice)
            if 1 <= idx <= len(excel_files):
                return excel_files[idx - 1].name
        matching = [archivo.name for archivo in excel_files if archivo.name.lower() == choice.lower()]
        if matching:
            return matching[0]
        print_colored('Entrada invalida. Intente nuevamente.', COLOR_RED)


def configure_cover_slide(presentation, lang):
    title_by_lang = {
        'P': 'Adicionais de Cobertura - 5W1H',
        'E': 'Adicionales de Cobertura - 5W1H'
    }
    cover_slide = presentation.slides[0]
    target_text = title_by_lang.get(lang, title_by_lang['E'])
    fallback_geometry = (Inches(0.42), Inches(2.0), Inches(10), Inches(0.8))
    geometry = None
    for shape in list(cover_slide.shapes):
        if not getattr(shape, 'has_text_frame', False):
            continue
        text_value = shape.text.strip()
        if 'Cobertura' in text_value and '5W1H' in text_value:
            geometry = (shape.left, shape.top, shape.width, shape.height)
            cover_slide.shapes._spTree.remove(shape._element)
    left, top, width, height = geometry if geometry else fallback_geometry
    text_box = cover_slide.shapes.add_textbox(left, top, width, height)
    tf = text_box.text_frame
    tf.clear()
    paragraph = tf.paragraphs[0]
    paragraph.text = target_text
    run = paragraph.runs[0]
    font = run.font
    font.name = 'Arial'
    font.size = Inches(0.45)
    font.bold = True
    font.color.rgb = RGBColor(255, 255, 255)


def plot_ven():
    options = [
        ("1 - Plotear Ventas y Compras Juntas", "1"),
        ("2 - Plotear Ventas y Compras Separadas", "2"),
        ("3 - No hay W con Ventas en esa base", "3"),
    ]

    print_colored('\nOpciones de grafico disponibles:', COLOR_QUESTION)
    for idx, (texto, _) in enumerate(options, 1):
        print_colored(f"  {idx}. {texto}", COLOR_BLUE)

    prompt = colorize(f"Seleccione el numero de la opcion (1-{len(options)}). Enter para {options[0][1]}: ", COLOR_QUESTION)
    while True:
        choice = input(prompt).strip()
        if not choice:
            return options[0][1]
        if choice.isdigit():
            idx = int(choice)
            if 1 <= idx <= len(options):
                return options[idx - 1][1]
        for _, valor in options:
            if choice == valor:
                return valor
        print_colored('Entrada invalida. Intente nuevamente.', COLOR_RED)

#Codigos paises
pais = pd.DataFrame(
    {'cod': [10, 54, 91, 55, 12, 56, 57, 93, 52, 51, 66, 63, 62, 64, 65, 67, 69],
     'pais': ['LatAm', 'Argentina', 'Bolivia', 'Brasil', 'CAM', 'Chile', 'Colombia', 'Ecuador', 'Mexico', 'Peru', 
              'Costa Rica', 'El Salvador', 'Guatemala', 'Honduras', 'Nicaragua', 'Panama', 'Republica Dominicana']
    }
)

#Codigos categorias
CATEG_CSV_DATA = """
cod,cest,cat
ALCB,Bebidas,Bebidas Alcohólicas
BEER,Bebidas,Cervezas
CARB,Bebidas,Bebidas Gaseosas
CWAT,Bebidas,Agua Gasificada
COCW,Bebidas,Água de Coco
COFF,Bebidas,Café-Consolidado de Café
CRBE,Bebidas,Cross Category (Bebidas)
ENDR,Bebidas,Bebidas Energéticas
FLBE,Bebidas,Bebidas Saborizadas Sin Gas
GCOF,Bebidas,Café Tostado y Molido
HJUI,Bebidas,Jugos Caseros
ITEA,Bebidas,Té Helado
ICOF,Bebidas,Café Instantáneo-Café Sucedáneo
JUNE,Bebidas,Jugos y Nectares
VEJU,Bebidas,Zumos de Vegetales
WATE,Bebidas,Agua Natural
CSDW,Bebidas,Gaseosas + Aguas
MXCM,Bebidas,Mixta Café+Malta
MXDG,Bebidas,Mixta Dolce Gusto-Mixta Té Helado + Café + Modificadores
MXJM,Bebidas,Mixta Jugos y Leches
MXJS,Bebidas,Mixta Jugos Líquidos + Bebidas de Soja
MXTC,Bebidas,Mixta Té+Café
JUIC,Bebidas,Jugos Liquidos-Jugos Polvo
PWDJ,Bebidas,Refrescos en Polvo-Jugos - Bebidas Instantáneas En Polvo - Jugos Polvo
RFDR,Bebidas,Bebidas Refrescantes
RTDJ,Bebidas,Refrescos Líquidos-Jugos Líquidos
RTEA,Bebidas,Té Líquido - Listo para Tomar
SOYB,Bebidas,Bebidas de Soja
SPDR,Bebidas,Bebidas Isotónicas
TEAA,Bebidas,Té e Infusiones-Te-Infusión Hierbas
YERB,Bebidas,Yerba Mate
BUTT,Lacteos,Manteca
CHEE,Lacteos,Queso Fresco y para Untar
CMLK,Lacteos,Leche Condensada
CRCH,Lacteos,Queso Untable
DYOG,Lacteos,Yoghurt p-beber
EMLK,Lacteos,Leche Culinaria-Leche Evaporada
FRMM,Lacteos,Leche Fermentada
FMLK,Lacteos,Leche Líquida Saborizada-Leche Líquida Con Sabor
FRMK,Lacteos,Fórmulas Infantiles
LQDM,Lacteos,Leche Líquida
LLFM,Lacteos,Leche Larga Vida
MARG,Lacteos,Margarina
MCHE,Lacteos,Queso Fundido
MKCR,Lacteos,Crema de Leche
MXDI,Lacteos,Mixta Lácteos-Postre+Leches+Yogurt
MXMI,Lacteos,Mixta Leches
MXYD,Lacteos,Mixta Yoghurt+Postres
PTSS,Lacteos,Petit Suisse
PWDM,Lacteos,Leche en Polvo
SYOG,Lacteos,Yoghurt p-comer
MILK,Lacteos,Leche-Leche Líquida Blanca - Leche Liq. Natural
YOGH,Lacteos,Yoghurt
CLOT,Ropas y Calzados,Ropas
FOOT,Ropas y Calzados,Calzados
SOCK,Ropas y Calzados,Medias-Calcetines
AREP,Alimentos,Arepas
BCER,Alimentos,Cereales Infantiles
BABF,Alimentos,Nutrición Infantil-Colados y Picados
BEAN,Alimentos,Frijoles
BISC,Alimentos,Galletas
BOUI,Alimentos,Caldos-Caldos y Sazonadores
BREA,Alimentos,Pan
BRCR,Alimentos,Apanados-Empanizadores
BRDC,Alimentos,Empanados
CERE,Alimentos,Cereales-Cereales Desayuno-Avenas y Cereales
BURG,Alimentos,Hamburguesas
CCMX,Alimentos,Mezclas Listas para Tortas-Preparados Base Harina Trigo
CAKE,Alimentos,Queques-Ponques Industrializados
FISH,Alimentos,Conservas De Pescado
CFAV,Alimentos,Conservas de Frutas y Verduras
CRML,Alimentos,Dulce de Leche-Manjar
CMLC,Alimentos,Alfajores
CBAR,Alimentos,Barras de Cereal
CHCK,Alimentos,Pollo
CHOC,Alimentos,Chocolate
COCO,Alimentos,Chocolate de Taza-Achocolatados - Cocoas
COLS,Alimentos,Salsas Frías
COMP,Alimentos,Compotas
SPIC,Alimentos,Condimentos y Especias
CKCH,Alimentos,Chocolate de Mesa
COIL,Alimentos,Aceite-Aceites Comestibles
CSAU,Alimentos,Salsas Listas-Salsas Caseras Envasadas
CNML,Alimentos,"Grano, Harina y Masa de Maíz"
CNST,Alimentos,Fécula de Maíz
CNFL,Alimentos,Harina De Maíz
CAID,Alimentos,Ayudantes Culinarios
DESS,Alimentos,Postres Preparados
DHAM,Alimentos,Jamón Endiablado
DFNS,Alimentos,Semillas y Frutos Secos
EBRE,Alimentos,Pan de Pascua
EEGG,Alimentos,Huevos de Páscua
EGGS,Alimentos,Huevos
FLSS,Alimentos,Flash Cecinas
FLOU,Alimentos,Harinas
MEAT,Alimentos,Carne Fresca
FRDS,Alimentos,Platos Listos Congelados
FRFO,Alimentos,Alimentos Congelados
HAMS,Alimentos,Jamones
HCER,Alimentos,Cereales Calientes-Cereales Precocidos
HOTS,Alimentos,Salsas Picantes
ICEC,Alimentos,Helados
IBRE,Alimentos,Pan Industrializado
IMPO,Alimentos,Puré Instantáneo
INOO,Alimentos,Fideos Instantáneos
JAMS,Alimentos,Mermeladas
KETC,Alimentos,Ketchup
LJDR,Alimentos,Jugo de Limon Adereso
MALT,Alimentos,Maltas
SEAS,Alimentos,Adobos - Sazonadores
MAYO,Alimentos,Mayonesa
MEAT,Alimentos,Cárnicos
MLKM,Alimentos,Modificadores de Leche-Saborizadores p-leche
MXCO,Alimentos,Mixta Cereales Infantiles+Avenas
MXBS,Alimentos,Mixta Caldos + Saborizantes
MXSB,Alimentos,Mixta Caldos + Sopas
MXCH,Alimentos,Mixta Cereales + Cereales Calientes
MXCC,Alimentos,Mixta Chocolate + Manjar
MXSN,Alimentos,"Galletas, snacks y mini tostadas"
COBT,Alimentos,Aceites + Mantecas
COCF,Alimentos,Aceites + Conservas De Pescado
CABB,Alimentos,Ayudantes Culinarios + Bolsa de Hornear
MXEC,Alimentos,Mixta Huevos de Páscua + Chocolates
MXDP,Alimentos,Mixta Platos Listos Congelados + Pasta
MXFR,Alimentos,Mixta Platos Congelados y Listos para Comer
MXFM,Alimentos,Mixta Alimentos Congelados + Margarina
MXMC,Alimentos,Mixta Modificadores + Cocoa
MXPS,Alimentos,Mixta Pastas
MXSO,Alimentos,Mixta Sopas+Cremas+Ramen
MXSP,Alimentos,Mixta Margarina + Mayonesa + Queso Crema
MXSW,Alimentos,Mixta Azúcar+Endulzantes
MUST,Alimentos,Mostaza
NDCR,Alimentos,Sustitutos de Crema
NOOD,Alimentos,Fideos
NUGG,Alimentos,Nuggets
OAFL,Alimentos,Avena en hojuelas-liquidas
OLIV,Alimentos,Aceitunas
PANC,Alimentos,Tortilla
PANE,Alimentos,Panetón
PAST,Alimentos,Pastas
PSAU,Alimentos,Salsas para Pasta
PNOU,Alimentos,Turrón de maní
PORK,Alimentos,Carne Porcina
PPMX,Alimentos,Postres en Polvo-Postres para Preparar - Horneables-Gelificables
PWSM,Alimentos,Leche de Soya en Polvo
PCCE,Alimentos,Cereales Precocidos
DOUG,Alimentos,Masas Frescas-Tapas Empanadas y Tarta
PPIZ,Alimentos,Pre-Pizzas
REFR,Alimentos,Meriendas listas
RICE,Alimentos,Arroz
RBIS,Alimentos,Galletas de Arroz
RTEB,Alimentos,Frijoles Procesados
RTEM,Alimentos,Pratos Prontos - Comidas Listas
SDRE,Alimentos,Aderezos para Ensalada
SALT,Alimentos,Sal
SLTC,Alimentos,Galletas Saladas-Galletas No Dulce
SARD,Alimentos,Sardina Envasada
SAUS,Alimentos,Cecinas
SCHN,Alimentos,Milanesas
SNAC,Alimentos,Snacks
SNOO,Alimentos,Fideos Sopa
SOUP,Alimentos,Sopas-Sopas Cremas
SOYS,Alimentos,Siyau
SPAG,Alimentos,Tallarines-Spaguetti
SPCH,Alimentos,Chocolate para Untar
SUGA,Alimentos,Azucar
SWCO,Alimentos,Galletas Dulces
SWSP,Alimentos,Untables Dulces
SWEE,Alimentos,Endulzantes
TOAS,Alimentos,Torradas - Tostadas
TOMA,Alimentos,Salsas de Tomate
TUNA,Alimentos,Atún Envasado
VMLK,Alimentos,Leche Vegetal
WFLO,Alimentos,Harinas de trigo
AIRC,Cuidado del Hogar,Ambientadores-Desodorante Ambiental
BARS,Cuidado del Hogar,Jabón en Barra-Jabón de lavar
BLEA,Cuidado del Hogar,Cloro-Lavandinas-Lejías-Blanqueadores
CBLK,Cuidado del Hogar,Pastillas para Inodoro
CGLO,Cuidado del Hogar,Guantes de látex
CLSP,Cuidado del Hogar,Esponjas de Limpieza-Esponjas y paños
CLTO,Cuidado del Hogar,Utensilios de Limpieza
FILT,Cuidado del Hogar,Filtros de Café
CRHC,Cuidado del Hogar,Cross Category (Limpiadores Domesticos)
CRLA,Cuidado del Hogar,Cross Category (Lavandería)
CRPA,Cuidado del Hogar,Cross Category (Productos de Papel)
DISH,Cuidado del Hogar,Lavavajillas-Lavaplatos - Lavalozas mano
DPAC,Cuidado del Hogar,Empaques domésticos-Bolsas plásticas-Plástico Adherente-Papel encerado-Papel aluminio
DRUB,Cuidado del Hogar,Destapacañerias
FBRF,Cuidado del Hogar,Perfumantes para Ropa-Perfumes para Ropa
FWAX,Cuidado del Hogar,Cera p-pisos
FDEO,Cuidado del Hogar,Desodorante para Pies
FRNP,Cuidado del Hogar,Lustramuebles
GBBG,Cuidado del Hogar,Bolsas de Basura
GCLE,Cuidado del Hogar,Limpiadores verdes
CLEA,Cuidado del Hogar,Limpiadores-Limpiadores y Desinfectantes
INSE,Cuidado del Hogar,Insecticidas-Raticidas
KITT,Cuidado del Hogar,Toallas de papel-Papel Toalla - Toallas de Cocina - Rollos Absorbentes de Papel
LAUN,Cuidado del Hogar,Detergentes para ropa
LSTA,Cuidado del Hogar,Apresto
MXBC,Cuidado del Hogar,Mixta Pastillas para Inodoro + Limpiadores
MXHC,Cuidado del Hogar,Mixta Home Care-Cloro-Limpiadores-Ceras-Ambientadores
MXCB,Cuidado del Hogar,Mixta Limpiadores + Cloro
MXLB,Cuidado del Hogar,Mixta Detergentes + Cloro
MXLD,Cuidado del Hogar,Mixta Detergentes + Lavavajillas
CRTO,Cuidado del Hogar,Pañitos + Papel Higienico
NAPK,Cuidado del Hogar,Servilletas
PLWF,Cuidado del Hogar,Film plastico e papel aluminio
SCOU,Cuidado del Hogar,Esponjas de Acero
SOFT,Cuidado del Hogar,Suavizantes de Ropa
STRM,Cuidado del Hogar,Quitamanchas-Desmanchadores
TOIP,Cuidado del Hogar,Papel Higiénico
WIPE,Cuidado del Hogar,Paños de Limpieza
ANLG,OTC,Analgésicos-Painkillers
FSUP,OTC,Suplementos alimentares
GMED,OTC,Gastrointestinales-Efervescentes
VITA,OTC,Vitaminas y Calcio
nan,Otros,Categoría Desconocida
BATT,Otros,Pilas-Baterías
CGAS,Otros,Combustible Gas
PFHH,Otros,Panel Financiero de Hogares
PFIN,Otros,Panel Financiero de Hogares
INKC,Otros,Cartuchos de Tintas
PETF,Otros,Alimento para Mascota-Alim.p - perro - gato
TELE,Otros,Telecomunicaciones - Convergencia
TILL,Otros,Tickets - Till Rolls
TOBA,Otros,Tabaco - Cigarrillos
ADIP,Cuidado Personal,Incontinencia de Adultos
BSHM,Cuidado Personal,Shampoo Infantil
RAZO,Cuidado Personal,Maquinas de Afeitar
BDCR,Cuidado Personal,Cremas Corporales
CWIP,Cuidado Personal,Paños Húmedos
COMB,Cuidado Personal,Cremas para Peinar
COND,Cuidado Personal,Acondicionador-Bálsamo
CRHY,Cuidado Personal,Cross Category (Higiene)
CRPC,Cuidado Personal,Cross Category (Personal Care)
DEOD,Cuidado Personal,Desodorantes
DIAP,Cuidado Personal,Pañales-Pañales Desechables
FCCR,Cuidado Personal,Cremas Faciales
FTIS,Cuidado Personal,Pañuelos Faciales
FEMI,Cuidado Personal,Protección Femenina-Toallas Femeninas
FRAG,Cuidado Personal,Fragancias
HAIR,Cuidado Personal,Cuidado del Cabello-Hair Care
HRCO,Cuidado Personal,Tintes para el Cabello-Tintes - Tintura - Tintes y Coloración para el cabello
HREM,Cuidado Personal,Depilación
HRST,Cuidado Personal,Alisadores para el Cabello
HSTY,Cuidado Personal,Fijadores para el Cabello-Modeladores-Gel-Fijadores para el cabello
HRTR,Cuidado Personal,Tratamientos para el Cabello
LINI,Cuidado Personal,Óleo Calcáreo
MAKE,Cuidado Personal,Maquillaje-Cosméticos
MEDS,Cuidado Personal,Jabón Medicinal
CRDT,Cuidado Personal,Pañitos + Pañales
MXMH,Cuidado Personal,Mixta Make Up+Tinturas
MOWA,Cuidado Personal,Enjuague Bucal-Refrescante Bucal
ORAL,Cuidado Personal,Cuidado Bucal
SPAD,Cuidado Personal,Protectores Femeninos
STOW,Cuidado Personal,Toallas Femininas
SHAM,Cuidado Personal,Shampoo
SHAV,Cuidado Personal,Afeitado-Crema afeitar-Loción de afeitar-Pord. Antes del afeitado
SKCR,Cuidado Personal,Cremas Faciales y Corporales-Cremas de Belleza - Cremas Cuerp y Faciales
SUNP,Cuidado Personal,Protección Solar
TALC,Cuidado Personal,Talcos-Talco para pies
TAMP,Cuidado Personal,Tampones Femeninos
TOIL,Cuidado Personal,Jabón de Tocador
TOOB,Cuidado Personal,Cepillos Dentales
TOOT,Cuidado Personal,Pastas Dentales
BAGS,Material Escolar,Morrales y MAletas Escoalres
CLPC,Material Escolar,Lapices de Colores
GRPC,Material Escolar,Lapices De Grafito
MRKR,Material Escolar,Marcadores
NTBK,Material Escolar,Cuadernos
SCHS,Material Escolar,Útiles Escolares
CSTD,Diversos,Estudio de Categorías
CORP,Diversos,Corporativa
CROS,Diversos,Cross Category
CRBA,Diversos,Cross Category (Bebés)
CRBR,Diversos,"Cross Category (Desayuno)-Yogurt, Cereal, Pan y Queso"
CRDT,Diversos,Cross Category (Diet y Light)
CRDF,Diversos,Cross Category (Alimentos Secos)
CRFO,Diversos,Cross Category (Alimentos)
CRSA,Diversos,Cross Category (Salsas)-Mayonesas-Ketchup - Salsas Frías
CRSN,Diversos,Cross Category (Snacks)
DEMO,Diversos,Demo
FLSH,Diversos,Flash
HLVW,Diversos,Holistic View
COCP,Diversos,Mezcla para café instantaneo y crema no láctea
CRSN,Diversos,Mezclas nutricionales y suplementos
MULT,Diversos,Consolidado-Multicategory
PCHK,Diversos,Pantry Check
STCK,Diversos,Inventario
MIHC,Diversos,Leche y Cereales Calientes-Cereales Precocidos y Leche Líquida Blanca
FLWT,Alimentos,Agua Saborizada
"""

categ = pd.read_csv(io.StringIO(CATEG_CSV_DATA), dtype={'cod': str, 'cest': str, 'cat': str})


#obtém o país,categoria,cesta e fabricante para template e ppt

base_dir = Path(__file__).resolve().parent

os.chdir(base_dir)

excel = select_excel_file(base_dir)



file = pd.ExcelFile(str(base_dir / excel))



W = file.sheet_names


#Obtém o pais cesta categoria fabricante marca e idioma para o qual se fará o estudo
land, cesta, cat, client = pais.loc[pais.cod==int(excel.split('_')[0]),'pais'].iloc[0], categ.loc[categ.cod==excel.split('_')[1],'cest'].iloc[0] ,categ.loc[categ.cod==excel.split('_')[1],'cat'].iloc[0] ,excel.split('_')[2].rsplit('.', 1)[0]

lang= "P" if land=='Brasil' else "E"

modelo_path = base_dir / 'Modelo_5W1H.pptx'

if not modelo_path.exists():

    raise FileNotFoundError(colorize(f'No se encontro el template {modelo_path.name} en {base_dir}', COLOR_RED))



ppt= Presentation(str(modelo_path))

configure_cover_slide(ppt, lang)



brand = W[0][2:]
#Dicionario com as correspondencias dos numeros e o respectivo W
c_w={
('P','1'):'1W - Quando?',
('P','3-1'):'3W - O quê? Tamanhos',
('P','3-2'):'3W - O quê? Marcas',
('P','3-3'):'3W - O quê? Sabores',
('P','4'):'4W - Quem? NSE',
('P','5-1'):'5W - Onde? Regiões',
('P','5-2'):'5W - Onde? Canais',
('P','6'):'Players',
('P','6-1'):'Players - Preco indexado',
('E','1'):'1W - ¿Cuándo?',
('E','3-1'):'3W - ¿Qué tamaños?',
('E','3-2'):'3W - ¿Qué marcas?',
('E','3-3'):'3W - ¿Qué sabores?',
('E','4'):'4W - ¿Quiénes? NSE (Nivel Socioeconómico)',
('E','5-1'):'5W - ¿Dónde? Regiones',
('E','5-2'):'5W - ¿Dónde? Canales',
('E','6'):'Players',
('E','6-1'):'Players - Precio indexado'
}

#Etiquetas dos slides
labels  ={
('P','Data'):'Data',
('P','MAT'):'Avaliação em Ano Móvel Acumulado',
('P','Var MAT'):"em ano móvel",
('P','comp'):"Concorrência do mercado de: ",

('E','Data'):'Fecha',
('E','MAT'):'Evaluación en Año Móvil Acumulado',
('E','Var MAT'):"en año móvil",
('E','comp'):"Competencia para el mercado de: "}

class SeriesConfig(NamedTuple):
    data: pd.DataFrame
    raw_tipo: str
    display_tipo: str
    pipeline: int


def _extract_pipeline(col_name: str) -> int:
    if not isinstance(col_name, str):
        return 0
    parts = [part for part in col_name.split('_') if part.isdigit()]
    if parts:
        return int(parts[0])
    digits = ''.join(ch for ch in col_name if ch.isdigit())
    return int(digits) if digits else 0


def _is_separator_column(col) -> bool:
    if col is None:
        return True
    if not isinstance(col, str):
        return False
    stripped = col.strip()
    return not stripped or stripped.lower().startswith('unnamed')


def split_compras_ventas(df: pd.DataFrame) -> tuple[list[pd.DataFrame], int]:
    ventas_idx = None
    for idx, col in enumerate(df.columns):
        if isinstance(col, str) and 'ventas' in col.lower():
            ventas_idx = idx
            pipeline = _extract_pipeline(col)
            break
    if ventas_idx is None:
        return [df], 0

    separator_idx = ventas_idx - 1
    has_separator = separator_idx >= 0 and _is_separator_column(df.columns[separator_idx])
    compras_end = separator_idx if has_separator else ventas_idx
    if compras_end <= 0:
        compras_end = ventas_idx

    compras = df.iloc[:, :compras_end].copy()
    ventas = df.iloc[:, ventas_idx:].copy()

    if not compras.empty:
        compras = compras.loc[:, ~pd.Series(compras.columns).apply(_is_separator_column).to_numpy()]
        if len(compras.columns) > 1:
            compras.columns = [compras.columns[0]] + [f"{col}.C" for col in compras.columns[1:]]

    if not ventas.empty:
        ventas = ventas.loc[:, ~pd.Series(ventas.columns).apply(_is_separator_column).to_numpy()]
        ventas_cols = list(ventas.columns)
        if ventas_cols:
            first = ventas_cols[0]
            base_name = first[:-2] if isinstance(first, str) and len(first) > 2 and first[-2] == '_' else first
            ventas_cols[0] = base_name
            ventas_cols[1:] = [str(col).replace('.1', '.V') for col in ventas_cols[1:]]
            ventas.columns = ventas_cols

    if ventas.shape[1] <= 1 or ventas.iloc[:, 1:].isna().all().all():
        return [compras if not compras.empty else df], 0

    df_list: list[pd.DataFrame] = []
    if not compras.empty:
        df_list.append(compras)
    df_list.append(ventas)
    return df_list, pipeline

def prepare_series_configs(df_list, lang, p_ventas):
    configs = []
    for original_df in df_list:
        df_local = original_df.copy()
        first_col = df_local.iloc[:,0]
        if first_col.dtype == object:
            df_local.iloc[:,0] = pd.to_datetime(first_col, format='%b-%y  ')
        raw_tipo = str(df_local.columns[0])
        df_local.rename(columns={df_local.columns[0]: labels[(lang,'Data')]}, inplace=True)
        is_compras = 'compras' in raw_tipo.lower()
        if not is_compras and len(df_local.columns) > 1:
            is_compras = '.c' in str(df_local.columns[1]).lower()
        display_tipo = 'Compras' if is_compras else 'Ventas'
        pipeline = 0 if is_compras else p_ventas
        configs.append(SeriesConfig(df_local, raw_tipo, display_tipo, pipeline))
    return configs


def extract_players_base_key(sheet_name: str) -> Optional[str]:
    if not isinstance(sheet_name, str):
        return None
    if not sheet_name or not sheet_name.startswith('6'):
        return None
    parts = [part for part in sheet_name.split('_') if part]
    if not parts:
        return None
    if parts[-1].isdigit():
        if len(parts) > 1:
            return '_'.join(parts[:-1])
        return parts[0]
    return '_'.join(parts)


def is_price_index_sheet(sheet_name: str) -> bool:
    if not isinstance(sheet_name, str):
        return False
    if not sheet_name.startswith('6'):
        return False
    parts = [part for part in sheet_name.split('_') if part]
    if len(parts) < 3:
        return False
    last_part = parts[-1]
    return last_part == '1'


def prepare_price_index_dataframe(series_config: SeriesConfig, top10_columns: list[str]) -> tuple[pd.DataFrame, list[str], pd.DataFrame]:
    df_price = series_config.data.copy()
    if df_price.empty or df_price.shape[1] <= 1:
        return pd.DataFrame(), [], pd.DataFrame()

    date_col = df_price.columns[0]
    data_cols = list(df_price.columns[1:])
    if not data_cols:
        return pd.DataFrame(), [], pd.DataFrame()

    total_col = next(
        (col for col in data_cols if 'total' in str(col).strip().lower()),
        data_cols[0]
    )

    df_price[total_col] = pd.to_numeric(df_price[total_col], errors='coerce')
    df_price = df_price[df_price[total_col].notna()]
    df_price = df_price[df_price[total_col] > 0]
    if df_price.empty:
        return pd.DataFrame(columns=[date_col]), [], pd.DataFrame(columns=[date_col])

    normalized_top10 = {str(col).strip().lower() for col in top10_columns if str(col).strip()}
    candidate_columns = []
    filtered_columns = []

    for col in data_cols:
        if col == total_col:
            continue
        numeric_series = pd.to_numeric(df_price[col], errors='coerce')
        if numeric_series.dropna().empty:
            continue
        if (numeric_series == 0).any():
            continue
        candidate_columns.append(col)
        col_lower = str(col).strip().lower()
        if not normalized_top10 or col_lower in normalized_top10:
            filtered_columns.append(col)

    if normalized_top10 and not filtered_columns:
        return pd.DataFrame(columns=[date_col]), [], pd.DataFrame(columns=[date_col])

    if not filtered_columns:
        filtered_columns = candidate_columns[:10]

    if not filtered_columns:
        return pd.DataFrame(columns=[date_col]), [], pd.DataFrame(columns=[date_col])

    total_numeric = pd.to_numeric(df_price[total_col], errors='coerce')
    ratio_df = pd.DataFrame({date_col: df_price[date_col]})
    ratio_df['Total'] = np.where(total_numeric.notna(), 100.0, np.nan)
    original_df = pd.DataFrame({date_col: df_price[date_col], 'Total': total_numeric})
    for col in filtered_columns:
        numeric_series = pd.to_numeric(df_price[col], errors='coerce')
        original_df[col] = numeric_series
        ratio_df[col] = numeric_series
    total_series = total_numeric.astype(float)
    with np.errstate(divide='ignore', invalid='ignore'):
        for col in filtered_columns:
            ratio_df[col] = ratio_df[col].astype(float) / total_series * 100.0
    ratio_df = ratio_df.replace([np.inf, -np.inf], np.nan)
    ratio_df.dropna(subset=filtered_columns, how='all', inplace=True)
    return ratio_df, ['Total'] + filtered_columns, original_df


def graf_price_index_table(
    columns: list[str],
    averages: dict[str, float],
    color_mapping: dict[str, str],
    metric_label: str,
    c_fig: int
) -> tuple[io.BytesIO, tuple[float, float]]:
    column_count = max(len(columns), 1)
    fig_width = max(6.0, 1.25 * column_count)
    fig_height = 1.5
    fig, ax = plt.subplots(num=c_fig, figsize=(fig_width, fig_height))
    ax.axis('off')
    ax.set_facecolor('white')

    header_labels = ['Indicador'] + [wrap_table_text(str(col)) for col in columns]
    display_columns = ['Indicador'] + columns
    row_values = [metric_label]
    for col in columns:
        value = averages.get(col)
        if value is None or not np.isfinite(value):
            row_values.append('-')
        else:
            row_values.append(f"{value:.2f}%")

    row_values[0] = wrap_table_text(row_values[0])

    table = ax.table(
        cellText=[row_values],
        colLabels=header_labels,
        cellLoc='center',
        loc='center'
    )

    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.05, 1.4)

    default_header_color = HEADER_COLOR_PRIMARY
    header_cell = table[(0, 0)]
    header_cell.set_facecolor(default_header_color)
    header_cell.set_text_props(color='white', weight='bold')
    header_cell.set_height(header_cell.get_height() * 1.5)
    for idx, col in enumerate(columns, start=1):
        header_cell = table[(0, idx)]
        header_color = color_mapping.get(col, '#777777')
        header_cell.set_facecolor(header_color)
        header_cell.set_text_props(color='white', weight='bold')
        header_cell.set_height(header_cell.get_height() * 1.5)

    for idx in range(len(display_columns)):
        data_cell = table[(1, idx)]
        data_cell.set_height(data_cell.get_height() * 1.35)
        data_cell.set_facecolor('#FFFFFF')
        data_cell.set_edgecolor('#DDDDDD')

    fig.tight_layout(pad=0.2)
    img_stream = io.BytesIO()
    fig.savefig(img_stream, format='png', transparent=True, dpi=220)
    img_stream.seek(0)
    plt.close(fig)
    return img_stream, fig.get_size_inches()


def build_price_index_slide(
    ppt: Presentation,
    lang: str,
    cat: str,
    apo_entries: list[dict],
    removed_headers: list[str],
    price_df: pd.DataFrame,
    price_original_df: pd.DataFrame,
    chart_pipeline: int,
    chart_share_lookup: dict,
    c_fig: int
) -> int:
    slide = ppt.slides.add_slide(ppt.slide_layouts[1])

    titulo_base = c_w.get((lang, '6-1'), 'Precio indexado')
    titulo_completo = f"{titulo_base} | {labels[(lang,'comp')]}{cat}"

    title_box = slide.shapes.add_textbox(Inches(0.33), Inches(0.2), Inches(10), Inches(0.5))
    title_tf = title_box.text_frame
    title_tf.clear()
    title_paragraph = title_tf.paragraphs[0]
    title_paragraph.text = titulo_completo
    title_paragraph.font.bold = True
    title_paragraph.font.size = Inches(0.33)

    comment_box = slide.shapes.add_textbox(Inches(11.07), Inches(6.33), Inches(2), Inches(0.5))
    comment_tf = comment_box.text_frame
    comment_tf.clear()
    comment_paragraph = comment_tf.paragraphs[0]
    comment_paragraph.text = "Comentário" if lang == 'P' else "Comentario"
    comment_paragraph.font.size = Inches(0.25)

    chart_share_lookup = chart_share_lookup or {}
    slide_height_in = emu_to_inches(ppt.slide_height)

    if not price_df.empty and price_df.shape[1] > 1:
        left_margin_chart = Inches(0.33)
        right_margin_chart = Inches(0.33)
        available_line_width = available_width(ppt, left_margin_chart, right_margin_chart)
        c_fig += 1
        chart_height = Cm(12)
        chart_colors = {}
        slide.shapes.add_picture(
            line_graf(
                price_df,
                chart_pipeline,
                titulo_completo,
                c_fig,
                1,
                width_emu=available_line_width,
                height_emu=chart_height,
                share_lookup=chart_share_lookup,
                y_axis_percent=True,
                top_value_annotations=2,
                color_collector=chart_colors,
                color_overrides={'Total': '#000000'},
                linestyle_overrides={'Total': '--'}
            ),
            left_margin_chart,
            Inches(1.15),
            width=available_line_width,
            height=chart_height
        )
        plt.clf()
        chart_colors.setdefault('Total', '#000000')

        if not price_original_df.empty and chart_colors:
            data_cols = [col for col in price_df.columns[1:] if col in chart_colors]
            if 'Total' in data_cols:
                data_cols = ['Total'] + [col for col in data_cols if col != 'Total']
            if data_cols:
                tail_df = price_original_df.iloc[-12:] if len(price_original_df) >= 12 else price_original_df
                averages_raw: dict[str, float] = {}
                for col in data_cols:
                    series = pd.to_numeric(tail_df[col], errors='coerce')
                    avg_value = series.mean() if not series.dropna().empty else np.nan
                    averages_raw[col] = avg_value

                total_key = next((col for col in data_cols if str(col).strip().lower() == 'total'), None)
                total_avg = averages_raw.get(total_key) if total_key else np.nan

                averages = {}
                if total_key and np.isfinite(total_avg) and not np.isclose(total_avg, 0.0):
                    for col, avg_value in averages_raw.items():
                        if avg_value is None or not np.isfinite(avg_value):
                            averages[col] = np.nan
                        elif col == total_key:
                            averages[col] = 100.0
                        else:
                            averages[col] = (avg_value / total_avg) * 100.0
                else:
                    for col, avg_value in averages_raw.items():
                        averages[col] = avg_value if avg_value is not None and np.isfinite(avg_value) else np.nan
                if any(np.isfinite(val) for val in averages.values()):
                    c_fig += 1
                    metric_label = "Preco indexado\nmedio 12 meses" if lang == 'P' else "Precio indexado 12m"
                    table_stream, fig_size = graf_price_index_table(
                        data_cols,
                        averages,
                        chart_colors,
                        metric_label,
                        c_fig
                    )
                    fig_width_in, fig_height_in = fig_size
                    if fig_width_in <= 0 or fig_height_in <= 0:
                        fig_width_in, fig_height_in = 6.0, 1.5
                    available_width_in = emu_to_inches(available_line_width)
                    if available_width_in <= 0:
                        available_width_in = fig_width_in
                    scale_ratio = available_width_in / fig_width_in
                    target_height_in = fig_height_in * scale_ratio
                    target_height_in = max(1.0, min(target_height_in, 2.2))
                    chart_top_in = 1.15
                    chart_height_in = emu_to_inches(chart_height)
                    table_gap_in = 0.2
                    table_top_in = chart_top_in + chart_height_in + table_gap_in
                    if table_top_in + target_height_in > slide_height_in - 0.3:
                        table_top_in = max(chart_top_in + chart_height_in + 0.05, slide_height_in - target_height_in - 0.3)
                    table_top = Inches(table_top_in)
                    table_width = available_line_width
                    table_height = Inches(target_height_in)
                    slide.shapes.add_picture(
                        table_stream,
                        left_margin_chart,
                        table_top,
                        width=table_width,
                        height=table_height
                    )
                    plt.clf()

    return c_fig








#Variavel de controle do numero de graficos
c_fig=0

plot=plot_ven()
last_reference_source = None
players_share_context = {}
#---------------------------------------------------------------------------------------------------------------------
for w in W:

    if is_price_index_sheet(w):
        players_base_key = extract_players_base_key(w)
        context = players_share_context.get(players_base_key)
        if context is None:
            print_colored(f"No se encontro contexto de Players para la hoja {w}. Se omite Precio indexado.", COLOR_YELLOW)
            continue

        df_start = file.parse(w)
        df_list, p_ventas = split_compras_ventas(df_start)
        series_configs = prepare_series_configs(df_list, lang, p_ventas)
        price_series = None
        for cfg in series_configs:
            if cfg.display_tipo.lower() == 'compras':
                price_series = cfg
                break
        if price_series is None:
            price_series = series_configs[0] if series_configs else None

        if price_series is None:
            print_colored(f"No se detecto informacion de Compras en la hoja {w}.", COLOR_YELLOW)
            continue

        price_df, selected_columns, price_original_df = prepare_price_index_dataframe(price_series, context.get('compras_top10', []))
        share_lookup_context = context.get('share_lookup', {})
        chart_share_lookup = {}
        if share_lookup_context and selected_columns:
            share_lookup_lower = {
                str(key).strip().lower(): value
                for key, value in share_lookup_context.items()
            }
            for col in selected_columns:
                share_val = share_lookup_lower.get(str(col).strip().lower())
                if share_val is not None:
                    chart_share_lookup[col] = share_val

        c_fig = build_price_index_slide(
            ppt,
            lang,
            cat,
            context.get('apo_entries', []),
            context.get('removed_headers', []),
            price_df,
            price_original_df,
            price_series.pipeline,
            chart_share_lookup,
            c_fig
        )
        last_reference_source = price_series.data
        titulo_precio = c_w.get((lang, '6-1'), 'Precio indexado')
        print_colored(f"{titulo_precio} realizado para {labels[(lang,'comp')]}{cat}", COLOR_GREEN)
        continue

    #1- Quando, grafico de variações MAT
    if w[0]=='1':

        #Cria o slide
        slide = ppt.slides.add_slide(ppt.slide_layouts[1])

        #Define o titulo
        txTitle = slide.shapes.add_textbox(Inches(0.33), Inches(0.2), Inches(10), Inches(0.5))
        tf = txTitle.text_frame
        tf.clear()
        t = tf.paragraphs[0]
        t.text = c_w[(lang,w[0])]+' '+ labels[(lang,'MAT')]+' | ' + w[2:] 
        t.font.bold = True
        t.font.size = Inches(0.35)

        #Obtém pipeline das vendas
        p=int(file.parse(w).columns[1].split("_")[1])

        #Cria a base
        mat=df_mat(file.parse(w),p)

        last_reference_source = mat
        
        #Elimina linhas com divisão com zero devido ao pipeline
        mat=mat[~np.isinf(mat.iloc[:, 3])]
      
        #Incrementa contador do gráfico
        c_fig+=1
        
        #Insere o gráfico do MAT
        left_margin = Inches(0.33)
        right_margin = Inches(0.33)
        available = available_width(ppt, left_margin, right_margin)
        pic=slide.shapes.add_picture(graf_mat(mat,c_fig,p), left_margin, Inches(1.15),width=available)

        #Insere caixa de texto para comentário do slide
        txTitle = slide.shapes.add_textbox(Inches(0.33), Inches(5.8), Inches(10), Inches(0.5))
        tf = txTitle.text_frame
        tf.clear()
        t = tf.paragraphs[0]
        t.text = "Comentário"
        t.font.size = Inches(0.28)

        #Limpa área de plotagem
        plt.clf()

        #mensaje de conclusion por cada slide
        print_colored(c_w[(lang,w[0])]+' realizado para '+ w[2:], COLOR_GREEN)

    #Outros
    else: 

        #Carrega a base
        df_start=file.parse(w)

        df_list, p_ventas = split_compras_ventas(df_start)

        series_configs = prepare_series_configs(df_list, lang, p_ventas)
        last_reference_source = series_configs[0].data if series_configs else df_start
        #Cria o slide
        slide = ppt.slides.add_slide(ppt.slide_layouts[1])

        #Define o titulo
        txTitle = slide.shapes.add_textbox(Inches(0.33), Inches(0.2), Inches(10), Inches(0.5))
        tf = txTitle.text_frame
        tf.clear()
        t = tf.paragraphs[0]
        if w[0] in ['3','5']:
            titulo = c_w[(lang,w[0]+'-'+w[-1])]+' | ' + w[2:-2] 
            t.text = titulo
            t.font.size = Inches(0.35)
        elif w[0]=='6':
            titulo =  c_w[(lang,w[0])] + ' | ' + labels[(lang,'comp')] + cat
            t.text =  titulo
            t.font.size = Inches(0.33)
        else:
            titulo = c_w[(lang,w[0])]+' | '+ w[2:] 
            t.text = titulo 
            t.font.size = Inches(0.35)

        t.font.bold = True

        #Insere caixa de texto para comentário do slide
        txTitle = slide.shapes.add_textbox(Inches(11.07), Inches(6.33), Inches(2), Inches(0.5))
        tf = txTitle.text_frame
        tf.clear()
        t = tf.paragraphs[0]
        t.text ="Comentário"
        t.font.size= Inches(0.25)

        #Controle posicao tabela aporte
        target_table_height = Cm(TABLE_TARGET_HEIGHT_CM)
        bottom_margin = Cm(1.5)
        left_margin = int(Cm(TABLE_SIDE_MARGIN_CM))
        right_margin = left_margin
        vertical_spacing = Cm(0.2)
        target_table_height_emu = int(target_table_height)
        bottom_margin_emu = int(bottom_margin)
        slide_width = int(ppt.slide_width)
        slide_height = int(ppt.slide_height)

        removed_headers = []
        series_share_lookup = {}
        stacked_share_sources = []
        should_collect_share = False
        apo_entries = []
        chart_color_mappings = {}
        if w:
            first_char = w[0]
            last_char = w[-1]
            if first_char == '6':
                should_collect_share = True
            elif first_char == '3' and last_char in {'1', '2'}:
                should_collect_share = True
            elif first_char == '5' and last_char == '2':
                should_collect_share = True

        for idx, serie in enumerate(series_configs):
            apo = aporte(serie.data.copy(), serie.pipeline, lang, serie.raw_tipo)
            apo_entries.append({
                "apo": apo,
                "display_tipo": serie.display_tipo,
            })
            removed_headers.extend(apo.attrs.get("removed_columns", []))
            share_values = apo.attrs.get("share_mat_values", {})
            if share_values:
                series_share_lookup.update(share_values)
            if should_collect_share:
                share_periods = apo.attrs.get("share_period_values", [])
                if share_periods:
                    stacked_share_sources.append(
                        {
                            "display_tipo": serie.display_tipo,
                            "share_periods": share_periods,
                            "columns": [
                                str(col)
                                for col in apo.columns[1:]
                                if str(col).strip() and str(col).strip().lower() != 'total'
                            ],
                        }
                    )

        #Cuadro de texto que indica los encabezados eliminados
        unique_removed_headers = [header for header in dict.fromkeys(removed_headers) if header]
        if unique_removed_headers:
            footer_left = Cm(4.6)
            footer_right_margin = Cm(1.2)
            footer_width = ppt.slide_width - footer_left - footer_right_margin
            if footer_width < Cm(2):
                footer_width = Cm(2)
            footer_top = ppt.slide_height - bottom_margin + Cm(0.1)
            footer_height = Cm(0.9)
            footer_box = slide.shapes.add_textbox(footer_left, footer_top, footer_width, footer_height)
            footer_tf = footer_box.text_frame
            footer_tf.clear()
            footer_tf.word_wrap = True
            footer_tf.text = "Se eliminaron los encabezados sin información: " + ", ".join(unique_removed_headers)
            footer_run = footer_tf.paragraphs[0].runs[0]
            footer_font = footer_run.font
            footer_font.name = 'Arial'
            footer_font.size = Pt(10)
            footer_font.color.rgb = RGBColor(120, 120, 120)

        compras_top10 = []
        if series_share_lookup:
            compras_items = []
            for col_name, share_value in series_share_lookup.items():
                if col_name is None:
                    continue
                col_str = str(col_name).strip()
                if not col_str:
                    continue
                if col_str.lower().find('total') != -1:
                    continue
                if not col_str.lower().endswith('.c'):
                    continue
                try:
                    numeric_share = float(share_value)
                except (TypeError, ValueError):
                    continue
                if not np.isfinite(numeric_share):
                    continue
                compras_items.append((col_name, numeric_share))
            if compras_items:
                compras_items.sort(key=lambda item: item[1], reverse=True)
                compras_top10 = [item[0] for item in compras_items[:10]]

        players_base_key = extract_players_base_key(w)
        if players_base_key and apo_entries:
            players_share_context[players_base_key] = {
                "apo_entries": apo_entries,
                "removed_headers": unique_removed_headers,
                "share_lookup": dict(series_share_lookup),
                "titulo": titulo,
                "compras_top10": compras_top10,
            }

        if plot=="1" and len(series_configs)>1:
            df_full = series_configs[0].data.copy()
            for extra in series_configs[1:]:
                df_full = pd.concat([df_full, extra.data.iloc[:,1:]], axis=1)
            pipeline_combined = max((cfg.pipeline for cfg in series_configs), default=0)
            c_fig+=1
            left_margin = Inches(0.33)
            right_margin = Inches(0.33)
            available_line_width = available_width(ppt, left_margin, right_margin)
            chart_colors = {}
            chart_stream = line_graf(
                df_full,
                pipeline_combined,
                titulo,
                c_fig,
                len(series_configs),
                width_emu=available_line_width,
                height_emu=Cm(10),
                share_lookup=series_share_lookup,
                color_collector=chart_colors
            )
            if isinstance(chart_colors, dict):
                chart_color_mappings.update(chart_colors)
            pic=slide.shapes.add_picture(chart_stream, left_margin, Inches(1.15),width=available_line_width,height=Cm(10))
        elif plot=="2" and len(series_configs)>1:
            ven_param = max(len(series_configs)-1, 1)
            left_margin = Inches(0.33)
            right_margin = Inches(0.33)
            available = available_width(ppt, left_margin, right_margin)
            count = len(series_configs)
            gap = Inches(0.1) if count > 1 else Inches(0)
            total_gap = int(gap) * (count - 1)
            effective_width = available - total_gap
            if count:
                if effective_width <= 0:
                    chart_width = available // count
                else:
                    chart_width = effective_width // count
            else:
                chart_width = available
            chart_width = max(chart_width, 1)
            for idx, serie in enumerate(series_configs):
                c_fig+=1
                left_position = left_margin + idx * (chart_width + int(gap))
                chart_colors = {}
                chart_stream = line_graf(
                    serie.data,
                    serie.pipeline,
                    titulo+' '+serie.display_tipo,
                    c_fig,
                    ven_param,
                    width_emu=chart_width,
                    height_emu=Cm(10),
                    multi_chart=True,
                    share_lookup=series_share_lookup,
                    color_collector=chart_colors
                )
                if isinstance(chart_colors, dict):
                    chart_color_mappings.update(chart_colors)
                pic=slide.shapes.add_picture(chart_stream, left_position, Inches(1.15),width=chart_width,height=Cm(10))
                plt.clf()
        elif series_configs:
            c_fig+=1
            left_margin = Inches(0.33)
            right_margin = Inches(0.33)
            available_single_width = available_width(ppt, left_margin, right_margin)
            chart_colors = {}
            chart_stream = line_graf(
                series_configs[0].data,
                series_configs[0].pipeline,
                titulo,
                c_fig,
                len(series_configs),
                width_emu=available_single_width,
                height_emu=Cm(10),
                share_lookup=series_share_lookup,
                color_collector=chart_colors
            )
            if isinstance(chart_colors, dict):
                chart_color_mappings.update(chart_colors)
            pic=slide.shapes.add_picture(chart_stream, left_margin, Inches(1.15),width=available_single_width,height=Cm(10))
            plt.clf()
        if apo_entries:
            normalized_colors = {
                str(key).strip(): value
                for key, value in chart_color_mappings.items()
                if key is not None and value
            }
            normalized_colors_lower = {
                key.lower(): value
                for key, value in normalized_colors.items()
                if key
            }
            suffix_options = ('.c', '.v', '_c', '_v', '-c', '-v')

            def column_color_for_table(column_name):
                if column_name is None:
                    return None
                key = str(column_name).strip()
                if not key or key.lower() == 'total':
                    return None
                color_value = normalized_colors.get(key)
                if color_value:
                    return color_value
                lower_key = key.lower()
                color_value = normalized_colors_lower.get(lower_key)
                if color_value:
                    return color_value
                for suffix in suffix_options:
                    if lower_key.endswith(suffix):
                        base = key[:-len(suffix)].strip()
                        if not base:
                            continue
                        color_value = normalized_colors.get(base)
                        if color_value:
                            return color_value
                        color_value = normalized_colors_lower.get(base.lower())
                        if color_value:
                            return color_value
                return None

            def build_table_color_mapping(apo_df):
                mapping = {}
                for column in apo_df.columns[1:]:
                    color_value = column_color_for_table(column)
                    if color_value:
                        mapping[column] = color_value
                return mapping

            table_top = slide_height - target_table_height_emu - bottom_margin_emu
            if len(apo_entries) == 2:
                gap = Cm(TABLE_PAIR_GAP_CM)
                half_slide_width = slide_width // 2
                gap_half = int(gap) // 2
                for entry in apo_entries:
                    apo_df = entry["apo"]
                    display_tipo = str(entry.get("display_tipo", "")).strip().lower()
                    if display_tipo == 'ventas':
                        left_position = half_slide_width + gap_half
                        max_width = slide_width - left_position - right_margin
                    else:
                        left_position = left_margin
                        max_width = half_slide_width - left_position - gap_half
                    c_fig += 1
                    table_colors = build_table_color_mapping(apo_df)
                    pic = slide.shapes.add_picture(
                        graf_apo(apo_df, c_fig, column_color_mapping=table_colors),
                        left_position,
                        table_top,
                        height=target_table_height
                    )
                    constrain_picture_width(pic, max_width)
                    plt.clf()
            else:
                for entry in apo_entries:
                    apo_df = entry["apo"]
                    c_fig += 1
                    table_colors = build_table_color_mapping(apo_df)
                    top_position = slide_height - target_table_height_emu - bottom_margin_emu
                    left_position = left_margin
                    max_width = slide_width - left_position - right_margin
                    pic = slide.shapes.add_picture(
                        graf_apo(apo_df, c_fig, column_color_mapping=table_colors),
                        left_position,
                        top_position,
                        height=target_table_height
                    )
                    constrain_picture_width(pic, max_width)
                    plt.clf()
        if should_collect_share and stacked_share_sources:
            priority_entry = next((entry for entry in stacked_share_sources if entry["display_tipo"].lower() == 'ventas'), stacked_share_sources[0])
            periods_to_plot = []
            for period_label, shares in priority_entry["share_periods"]:
                if shares and any(value > 0 for value in shares.values()):
                    periods_to_plot.append((period_label, shares))
                if len(periods_to_plot) == 2:
                    break
            if periods_to_plot:
                previous_shares = periods_to_plot[0][1] if len(periods_to_plot) > 1 else None
                delta_header = 'Crescimento vs ano anterior' if lang == 'P' else 'Crecimiento vs año anterior'
                share_slide = ppt.slides.add_slide(ppt.slide_layouts[1])
                share_title_box = share_slide.shapes.add_textbox(Inches(0.33), Inches(0.2), Inches(10), Inches(0.5))
                share_tf = share_title_box.text_frame
                share_tf.clear()
                share_title_paragraph = share_tf.paragraphs[0]
                share_title_paragraph.text = f"{titulo} - Share 100% Apilado"
                share_title_paragraph.font.bold = True
                share_title_paragraph.font.size = Inches(0.33)

                comment_box = share_slide.shapes.add_textbox(Inches(11.07), Inches(6.33), Inches(2), Inches(0.5))
                comment_tf = comment_box.text_frame
                comment_tf.clear()
                comment_paragraph = comment_tf.paragraphs[0]
                comment_paragraph.text = "Comentário"
                comment_paragraph.font.size = Inches(0.25)

                palette_values = TREND_COLOR_SEQUENCE or [mcolors.to_hex(c) for c in plt.get_cmap('tab20').colors]
                color_mapping = {}
                column_candidates = [col for col in priority_entry["columns"] if col]
                if not column_candidates:
                    for _, shares in periods_to_plot:
                        for column_name in shares.keys():
                            if column_name and column_name not in column_candidates:
                                column_candidates.append(column_name)
                for idx_color, column_name in enumerate(column_candidates):
                    color_mapping[column_name] = palette_values[idx_color % len(palette_values)]

                chart_top = Inches(0.55)
                left_start = Inches(2.0)
                right_margin = Inches(0.6)
                horizontal_gap = Inches(0.5)
                caption_gap = Inches(0.12)
                caption_height = Inches(0.35)
                current_left = left_start

                reference_label = 'Corte em' if lang == 'P' else 'Corte a'

                for idx_chart, (period_label, shares) in enumerate(periods_to_plot[:2]):
                    total_share = sum(max(value, 0) for value in shares.values())
                    if total_share <= 0:
                        continue
                    c_fig += 1
                    chart_stream, fig_size = stacked_share_chart(
                        period_label,
                        shares,
                        color_mapping,
                        c_fig,
                        title=f"{priority_entry['display_tipo']} - {period_label}"
                    )
                    fig_width_in, fig_height_in = fig_size
                    if fig_height_in <= 0:
                        continue
                    aspect_ratio = fig_width_in / fig_height_in if fig_height_in else 1.0
                    target_height_in = max(6.0, fig_height_in * 1.05)
                    slide_height_in = emu_to_inches(ppt.slide_height)
                    chart_top_in = emu_to_inches(chart_top)
                    caption_gap_in = emu_to_inches(caption_gap)
                    caption_height_in = emu_to_inches(caption_height)
                    available_height_in = max(3.0, slide_height_in - chart_top_in - caption_gap_in - caption_height_in - 0.2)
                    target_height_in = min(available_height_in, target_height_in)
                    target_width_in = max(3.8, target_height_in * aspect_ratio)
                    slide_width_in = emu_to_inches(ppt.slide_width)
                    available_width_in = slide_width_in - emu_to_inches(current_left) - emu_to_inches(right_margin)
                    if available_width_in <= 0:
                        continue
                    if target_width_in > available_width_in:
                        target_width_in = available_width_in
                        target_height_in = target_width_in / aspect_ratio if aspect_ratio else target_height_in
                    target_height = Inches(target_height_in)
                    target_width = Inches(target_width_in)
                    chart_shape = share_slide.shapes.add_picture(
                        chart_stream,
                        current_left,
                        chart_top,
                        width=target_width,
                        height=target_height
                    )
                    caption_left = chart_shape.left
                    caption_width = target_width
                    caption_top = chart_top + target_height + caption_gap
                    caption_box = share_slide.shapes.add_textbox(caption_left, caption_top, caption_width, caption_height)
                    caption_tf = caption_box.text_frame
                    caption_tf.clear()
                    caption_paragraph = caption_tf.paragraphs[0]
                    caption_paragraph.text = f"{reference_label}: {period_label}"
                    caption_paragraph.font.size = Pt(11)
                    caption_paragraph.font.bold = True
                    caption_paragraph.font.color.rgb = RGBColor(80, 80, 80)

                    if previous_shares is not None and idx_chart == len(periods_to_plot[:2]) - 1:
                        delta_box_width = Inches(2.6)
                        delta_box_left = chart_shape.left + chart_shape.width + Inches(0.2)
                        slide_right_limit = ppt.slide_width - Inches(0.2)
                        if delta_box_left + delta_box_width > slide_right_limit:
                            adjusted_left = slide_right_limit - delta_box_width
                            if adjusted_left < Inches(0.4):
                                delta_box_left = max(chart_shape.left - delta_box_width - Inches(0.2), Inches(0.4))
                            else:
                                delta_box_left = adjusted_left
                        delta_box = share_slide.shapes.add_textbox(
                            delta_box_left,
                            chart_top,
                            delta_box_width,
                            target_height
                        )
                        delta_tf = delta_box.text_frame
                        delta_tf.clear()
                        delta_tf.word_wrap = True
                        header_para = delta_tf.paragraphs[0]
                        header_para.text = delta_header
                        header_para.font.bold = True
                        header_para.font.size = Pt(12)
                        header_para.font.color.rgb = RGBColor(70, 70, 70)

                        growth_entries = []
                        for brand, current_value in shares.items():
                            prev_value = previous_shares.get(brand, 0.0) if previous_shares else 0.0
                            if current_value is None or not np.isfinite(current_value):
                                current_value = 0.0
                            if prev_value is None or not np.isfinite(prev_value):
                                prev_value = 0.0
                            delta_value = (current_value - prev_value) * 100
                            growth_entries.append((brand, delta_value, current_value))
                        growth_entries.sort(key=lambda item: item[1], reverse=True)

                        for brand, delta_value, _ in growth_entries:
                            para = delta_tf.add_paragraph()
                            para.level = 1
                            para.text = f"{brand}: {delta_value:+.1f} pp"
                            para.font.size = Pt(11)
                            color_hex = color_mapping.get(brand)
                            if color_hex:
                                rgb_tuple = mcolors.to_rgb(color_hex)
                                color_rgb = tuple(int(round(c * 255)) for c in rgb_tuple)
                                para.font.color.rgb = RGBColor(*color_rgb)
                            else:
                                para.font.color.rgb = RGBColor(90, 90, 90)

                    current_left += target_width + horizontal_gap
                    plt.clf()
        if w[0] in ['3','5']:
            print_colored(c_w[(lang,w[0]+'-'+w[-1])]+' realizado para ' + w[2:-2], COLOR_GREEN) 
        elif w[0]=='6':
            print_colored(c_w[(lang,w[0])] + ' realizado para ' + labels[(lang,'comp')] + cat, COLOR_GREEN)
        else:
            print_colored(c_w[(lang,w[0])]+' realizado para '+ w[2:], COLOR_GREEN)


#Referencia da base
ref_source = last_reference_source if last_reference_source is not None else file.parse(W[0])
last_value = ref_source.iloc[-1, 0] if not ref_source.empty else None
if last_value is None:
    ref = 'NA'
    ref_display = 'NA'
else:
    if isinstance(last_value, str):
        try:
            last_dt = dt.strptime(last_value, '%b-%y  ')
        except ValueError:
            last_dt = pd.to_datetime(last_value).to_pydatetime()
    else:
        last_dt = pd.to_datetime(last_value).to_pydatetime()
    ref = last_dt.strftime('%m-%y')
    month_list = MONTH_NAMES.get(lang, MONTH_NAMES['E'])
    month_name = month_list[last_dt.month - 1]
    ref_display = f"{month_name}-{last_dt.year}"
slide=ppt.slides[0]

#Fabricante
txBox = slide.shapes.add_textbox(Inches(0.42), Inches(3.77), Inches(10), Inches(0.5))
tf = txBox.text_frame
tf.clear()
t = tf.paragraphs[0]
t.text = "Numerator vs " + client
run = t.runs[0]
font = run.font
font.name, font.size, font.bold = 'Arial', Inches(0.44), True
font.color.rgb = RGBColor(255, 255, 255)

#Referencia
txBox = slide.shapes.add_textbox(Inches(0.42), Inches(5.98), Inches(10), Inches(0.5))
tf = txBox.text_frame
tf.clear()
t = tf.paragraphs[0]
reference_label = 'Corte em' if lang == 'P' else 'Corte a'
display_ref = ref_display if ref_display != 'NA' else ref
t.text = f"{reference_label} {display_ref}"
run = t.runs[0]
font = run.font
font.name, font.size, font.bold = 'Arial', Inches(0.32), True
font.color.rgb = RGBColor(255, 255, 255)

output_filename = "-".join([
    simplify_name_segment(land, 30),
    simplify_name_segment(client, 30),
    simplify_name_segment(cat, 6),
    simplify_name_segment(brand, 8),
    '5W1H',
    simplify_name_segment(ref, 5)
]) + '.pptx'

ppt.save(output_filename)

fim = dt.now()

t = int((fim - agora).total_seconds())
print_colored(f'Tiempo de ejecucion : {t//60} min {t%60} s' if t >= 60 else f'Tiempo de ejecucion : {t} s', COLOR_BLUE)
