# -*- coding: utf-8 -*-
#TODO: Solucionar las tablas dinamicas. que el ancho se adapte al contenido. si es muy grande que se adapte al slide, y si es muy chica que no ocupe todo el slide

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
from typing import NamedTuple
from pathlib import Path

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


EMU_PER_INCH = 914400
DEFAULT_LINE_CHART_RATIO = 3
TABLE_TARGET_HEIGHT_CM = 4.0
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
def line_graf(df, p, title, c_fig, ven, width_emu=None, height_emu=None, multi_chart=None):
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
            color_mapping[base] = palette_values[color_index % len(palette_values)]
            color_index += 1

    lns = []
    legend_labels = []
    numeric_series_list = []
    for col in colunas:
        if ven > 1:
            estilo = '-' if '.v' in col.lower() else '--'
        else:
            estilo = '-'

        cor = color_mapping.get(col, palette_values[0])

        numeric_series = pd.to_numeric(aux[col], errors='coerce')
        numeric_series_list.append(numeric_series)
        y = numeric_series.values
        x_slice = x_positions[start_idx:]
        y_slice = y[start_idx:]
        valid_points = [(x_idx, y_val) for x_idx, y_val in zip(x_slice, y_slice) if pd.notna(y_val)]
        if not valid_points:
            continue

        suffix_map = {".c": "Compras", ".v": "Ventas", "_c": "Compras", "_v": "Ventas", "-c": "Compras", "-v": "Ventas"}
        lower_col = col.lower()
        legend_label = col
        for suffix, suffix_text in suffix_map.items():
            if lower_col.endswith(suffix):
                legend_label = f"{col[:-len(suffix)].strip()} ({suffix_text})"
                break

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


    if data_len and start_idx < data_len:
        ax.set_xticks(x_positions[start_idx:])
        ax.set_xticklabels(ran[start_idx:], rotation=30, fontsize=xtick_size)
        ax.set_xlim(x_positions[start_idx], x_positions[-1])
    else:
        ax.set_xticks([])
        ax.set_xticklabels([])

    y_tick_size = max(8, int(10 * scale))
    ax.tick_params(axis='y', labelsize=y_tick_size)

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

    img_stream = io.BytesIO()
    fig.savefig(img_stream, format='png', transparent=True)
    img_stream.seek(0)

    try:
        plt.close(fig)
    except Exception:
        pass

    return img_stream

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

        #Formatação do volume
        apo.iloc[:2, 1:] = apo.iloc[:2, 1:].applymap(lambda x: f"{round(x * 100, 1)}%")
        #Formatação da variação e aporte
        apo.iloc[2:, 1:] = apo.iloc[2:, 1:].applymap(lambda x: f"{round(x * 100, 2)}%")

        return apo

#Função que cria o gráfico tabela de aporte

def graf_apo(apo,c_fig):
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

    column_colors = share_colors(len(data_columns))
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
            text.set_color(HEADER_FONT_COLOR)
            if col == 0 or (total_column_name is not None and apo.columns[col] == total_column_name):
                cell.set_facecolor(header_main)
            else:
                cell.set_facecolor(header_secondary)
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
categ = pd.DataFrame({
    'cest': ['Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 
               'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Bebidas', 'Lacteos', 
               'Lacteos', 'Lacteos', 'Lacteos', 'Lacteos', 'Lacteos', 'Lacteos', 'Lacteos', 'Lacteos', 'Lacteos', 'Lacteos', 'Lacteos', 'Lacteos', 'Lacteos', 'Lacteos', 'Lacteos', 'Lacteos', 
               'Lacteos', 'Lacteos', 'Lacteos', 'Lacteos', 'Lacteos', 'Ropas y Calzados', 'Ropas y Calzados', 'Ropas y Calzados', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 
               'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 
               'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 
               'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 
               'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 
               'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 
               'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 
               'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 
               'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 'Alimentos', 
               'Alimentos', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 
               'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 
               'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 
               'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 
               'Cuidado del Hogar', 'Cuidado del Hogar', 'Cuidado del Hogar', 'OTC', 'OTC', 'OTC', 'OTC', 'Otros', 'Otros', 'Otros', 'Otros', 'Otros', 'Otros', 'Otros', 'Otros', 'Otros', 'Otros', 
               'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 
               'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 
               'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 
               'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 
               'Cuidado Personal', 'Cuidado Personal', 'Cuidado Personal', 'Material Escolar', 'Material Escolar', 'Material Escolar', 'Material Escolar', 'Material Escolar', 'Material Escolar', 
               'Diversos', 'Diversos', 'Diversos', 'Diversos', 'Diversos', 'Diversos', 'Diversos', 'Diversos', 'Diversos', 'Diversos', 'Diversos', 'Diversos', 'Diversos', 'Diversos', 'Diversos', 
               'Diversos', 'Diversos', 'Diversos', 'Diversos','Alimentos'],
    'cat': ['Bebidas Alcohólicas', 'Cervezas', 'Bebidas Gaseosas', 'Agua Gasificada', 'Água de Coco', 'Café-Consolidado de Café', 'Cross Category (Bebidas)', 'Bebidas Energéticas', 
                  'Bebidas Saborizadas Sin Gas', 'Café Tostado y Molido', 'Jugos Caseros', 'Té Helado', 'Café Instantáneo-Café Sucedáneo', 'Jugos y Nectares', 'Zumos de Vegetales', 'Agua Natural', 
                  'Gaseosas + Aguas', 'Mixta Café+Malta', 'Mixta Dolce Gusto-Mixta Té Helado + Café + Modificadores', 'Mixta Jugos y Leches', 'Mixta Jugos Líquidos + Bebidas de Soja', 'Mixta Té+Café', 
                  'Jugos Liquidos-Jugos Polvo', 'Refrescos en Polvo-Jugos - Bebidas Instantáneas En Polvo - Jugos Polvo', 'Bebidas Refrescantes', 'Refrescos Líquidos-Jugos Líquidos', 'Té Líquido - Listo para Tomar',
                  'Bebidas de Soja', 'Bebidas Isotónicas', 'Té e Infusiones-Te-Infusión Hierbas', 'Yerba Mate', 'Manteca', 'Queso Fresco y para Untar', 'Leche Condensada', 'Queso Untable', 'Yoghurt p-beber', 
                  'Leche Culinaria-Leche Evaporada', 'Leche Fermentada', 'Leche Líquida Saborizada-Leche Líquida Con Sabor', 'Fórmulas Infantiles', 'Leche Líquida', 'Leche Larga Vida', 'Margarina', 'Queso Fundido', 
                  'Crema de Leche', 'Mixta Lácteos-Postre+Leches+Yogurt', 'Mixta Leches', 'Mixta Yoghurt+Postres', 'Petit Suisse', 'Leche en Polvo', 'Yoghurt p-comer', 'Leche-Leche Líquida Blanca - Leche Liq. Natural', 
                  'Yoghurt', 'Ropas', 'Calzados', 'Medias-Calcetines', 'Arepas', 'Cereales Infantiles', 'Nutrición Infantil-Colados y Picados', 'Frijoles', 'Galletas', 'Caldos-Caldos y Sazonadores', 'Pan', 
                  'Apanados-Empanizadores', 'Empanados', 'Cereales-Cereales Desayuno-Avenas y Cereales', 'Hamburguesas', 'Mezclas Listas para Tortas-Preparados Base Harina Trigo', 'Queques-Ponques Industrializados', 
                  'Conservas De Pescado', 'Conservas de Frutas y Verduras', 'Dulce de Leche-Manjar', 'Alfajores', 'Barras de Cereal', 'Pollo', 'Chocolate', 'Chocolate de Taza-Achocolatados - Cocoas', 'Salsas Frías', 
                  'Compotas', 'Condimentos y Especias', 'Chocolate de Mesa', 'Aceite-Aceites Comestibles', 'Salsas Listas-Salsas Caseras Envasadas', 'Grano, Harina y Masa de Maíz', 'Fécula de Maíz', 'Harina De Maíz', 
                  'Ayudantes Culinarios', 'Postres Preparados', 'Jamón Endiablado', 'Semillas y Frutos Secos', 'Pan de Pascua', 'Huevos de Páscua', 'Huevos', 'Flash Cecinas', 'Harinas', 'Carne Fresca', 
                  'Platos Listos Congelados','Alimentos Congelados', 'Jamones', 'Cereales Calientes-Cereales Precocidos', 'Salsas Picantes', 'Helados', 'Pan Industrializado', 'Puré Instantáneo', 
                  'Fideos Instantáneos', 'Mermeladas', 'Ketchup', 'Jugo de Limon Adereso', 'Maltas', 'Adobos - Sazonadores', 'Mayonesa', 'Cárnicos', 'Modificadores de Leche-Saborizadores p-leche', 
                  'Mixta Cereales Infantiles+Avenas', 'Mixta Caldos + Saborizantes', 'Mixta Caldos + Sopas', 'Mixta Cereales + Cereales Calientes', 'Mixta Chocolate + Manjar', 'Galletas, snacks y mini tostadas', 
                  'Aceites + Mantecas', 'Aceites + Conservas De Pescado', 'Ayudantes Culinarios + Bolsa de Hornear', 'Mixta Huevos de Páscua + Chocolates', 'Mixta Platos Listos Congelados + Pasta', 
                  'Mixta Platos Congelados y Listos para Comer', 'Mixta Alimentos Congelados + Margarina', 'Mixta Modificadores + Cocoa', 'Mixta Pastas', 'Mixta Sopas+Cremas+Ramen', 
                  'Mixta Margarina + Mayonesa + Queso Crema', 'Mixta Azúcar+Endulzantes', 'Mostaza', 'Sustitutos de Crema', 'Fideos', 'Nuggets', 'Avena en hojuelas-liquidas', 'Aceitunas', 'Tortilla', 
                  'Panetón', 'Pastas', 'Salsas para Pasta', 'Turrón de maní', 'Carne Porcina', 'Postres en Polvo-Postres para Preparar - Horneables-Gelificables', 'Leche de Soya en Polvo', 'Cereales Precocidos', 
                  'Masas Frescas-Tapas Empanadas y Tarta', 'Pre-Pizzas', 'Meriendas listas', 'Arroz', 'Galletas de Arroz', 'Frijoles Procesados', 'Pratos Prontos - Comidas Listas', 'Aderezos para Ensalada', 
                  'Sal', 'Galletas Saladas-Galletas No Dulce', 'Sardina Envasada', 'Cecinas', 'Milanesas', 'Snacks', 'Fideos Sopa', 'Sopas-Sopas Cremas', 'Siyau', 'Tallarines-Spaguetti', 'Chocolate para Untar', 
                  'Azucar', 'Galletas Dulces', 'Untables Dulces', 'Endulzantes', 'Torradas - Tostadas', 'Salsas de Tomate', 'Atún Envasado', 'Leche Vegetal', 'Harinas de trigo', 
                  'Ambientadores-Desodorante Ambiental', 'Jabón en Barra-Jabón de lavar', 'Cloro-Lavandinas-Lejías-Blanqueadores', 'Pastillas para Inodoro', 'Guantes de látex', 
                  'Esponjas de Limpieza-Esponjas y paños', 'Utensilios de Limpieza', 'Filtros de Café', 'Cross Category (Limpiadores Domesticos)', 'Cross Category (Lavandería)', 
                  'Cross Category (Productos de Papel)', 'Lavavajillas-Lavaplatos - Lavalozas mano', 'Empaques domésticos-Bolsas plásticas-Plástico Adherente-Papel encerado-Papel aluminio', 
                  'Destapacañerias', 'Perfumantes para Ropa-Perfumes para Ropa', 'Cera p-pisos', 'Desodorante para Pies', 'Lustramuebles', 'Bolsas de Basura', 'Limpiadores verdes', 
                  'Limpiadores-Limpiadores y Desinfectantes', 'Insecticidas-Raticidas', 'Toallas de papel-Papel Toalla - Toallas de Cocina - Rollos Absorbentes de Papel', 'Detergentes para ropa', 
                  'Apresto', 'Mixta Pastillas para Inodoro + Limpiadores', 'Mixta Home Care-Cloro-Limpiadores-Ceras-Ambientadores', 'Mixta Limpiadores + Cloro', 'Mixta Detergentes + Cloro', 
                  'Mixta Detergentes + Lavavajillas', 'Pañitos + Papel Higienico', 'Servilletas', 'Film plastico e papel aluminio', 'Esponjas de Acero', 'Suavizantes de Ropa', 'Quitamanchas-Desmanchadores', 
                  'Papel Higiénico', 'Paños de Limpieza', 'Analgésicos-Painkillers', 'Suplementos alimentares', 'Gastrointestinales-Efervescentes', 'Vitaminas y Calcio', 'Categoría Desconocida', 
                  'Pilas-Baterías', 'Combustible Gas', 'Panel Financiero de Hogares', 'Panel Financiero de Hogares', 'Cartuchos de Tintas', 'Alimento para Mascota-Alim.p - perro - gato', 
                  'Telecomunicaciones - Convergencia', 'Tickets - Till Rolls', 'Tabaco - Cigarrillos', 'Incontinencia de Adultos', 'Shampoo Infantil', 'Maquinas de Afeitar', 'Cremas Corporales', 
                  'Paños Húmedos', 'Cremas para Peinar', 'Acondicionador-Bálsamo', 'Cross Category (Higiene)', 'Cross Category (Personal Care)', 'Desodorantes', 'Pañales-Pañales Desechables', 
                  'Cremas Faciales', 'Pañuelos Faciales', 'Protección Femenina-Toallas Femeninas', 'Fragancias', 'Cuidado del Cabello-Hair Care', 
                  'Tintes para el Cabello-Tintes - Tintura - Tintes y Coloración para el cabello', 'Depilación', 'Alisadores para el Cabello', 
                  'Fijadores para el Cabello-Modeladores-Gel-Fijadores para el cabello', 'Tratamientos para el Cabello', 'Óleo Calcáreo', 'Maquillaje-Cosméticos', 'Jabón Medicinal', 
                  'Pañitos + Pañales', 'Mixta Make Up+Tinturas', 'Enjuague Bucal-Refrescante Bucal', 'Cuidado Bucal', 'Protectores Femeninos', 'Toallas Femininas', 'Shampoo', 
                  'Afeitado-Crema afeitar-Loción de afeitar-Pord. Antes del afeitado', 'Cremas Faciales y Corporales-Cremas de Belleza - Cremas Cuerp y Faciales', 'Protección Solar', 
                  'Talcos-Talco para pies', 'Tampones Femeninos', 'Jabón de Tocador', 'Cepillos Dentales', 'Pastas Dentales', 'Morrales y MAletas Escoalres', 'Lapices de Colores', 'Lapices De Grafito', 
                  'Marcadores', 'Cuadernos', 'Útiles Escolares', 'Estudio de Categorías', 'Corporativa', 'Cross Category', 'Cross Category (Bebés)', 'Cross Category (Desayuno)-Yogurt, Cereal, Pan y Queso', 
                  'Cross Category (Diet y Light)', 'Cross Category (Alimentos Secos)', 'Cross Category (Alimentos)', 'Cross Category (Salsas)-Mayonesas-Ketchup - Salsas Frías', 'Cross Category (Snacks)', 
                  'Demo', 'Flash', 'Holistic View', 'Mezcla para café instantaneo y crema no láctea', 'Mezclas nutricionales y suplementos', 'Consolidado-Multicategory', 'Pantry Check', 'Inventario', 
                  'Leche y Cereales Calientes-Cereales Precocidos y Leche Líquida Blanca','Agua Saborizada'],
    'cod': ['ALCB', 'BEER', 'CARB', 'CWAT', 'COCW', 'COFF', 'CRBE', 'ENDR', 'FLBE', 'GCOF', 'HJUI', 'ITEA', 'ICOF', 'JUNE', 'VEJU', 'WATE', 'CSDW', 'MXCM', 'MXDG', 'MXJM', 'MXJS', 'MXTC', 'JUIC', 'PWDJ', 
               'RFDR', 'RTDJ', 'RTEA', 'SOYB', 'SPDR', 'TEAA', 'YERB', 'BUTT', 'CHEE', 'CMLK', 'CRCH', 'DYOG', 'EMLK', 'FRMM', 'FMLK', 'FRMK', 'LQDM', 'LLFM', 'MARG', 'MCHE', 'MKCR', 'MXDI', 'MXMI', 'MXYD', 
               'PTSS', 'PWDM', 'SYOG', 'MILK', 'YOGH', 'CLOT', 'FOOT', 'SOCK', 'AREP', 'BCER', 'BABF', 'BEAN', 'BISC', 'BOUI', 'BREA', 'BRCR', 'BRDC', 'CERE', 'BURG', 'CCMX', 'CAKE', 'FISH', 'CFAV', 'CRML', 
               'CMLC', 'CBAR', 'CHCK', 'CHOC', 'COCO', 'COLS', 'COMP', 'SPIC', 'CKCH', 'COIL', 'CSAU', 'CNML', 'CNST', 'CNFL', 'CAID', 'DESS', 'DHAM', 'DFNS', 'EBRE', 'EEGG', 'EGGS', 'FLSS', 'FLOU', 'MEAT', 
               'FRDS', 'FRFO', 'HAMS', 'HCER', 'HOTS', 'ICEC', 'IBRE', 'IMPO', 'INOO', 'JAMS', 'KETC', 'LJDR', 'MALT', 'SEAS', 'MAYO', 'MEAT', 'MLKM', 'MXCO', 'MXBS', 'MXSB', 'MXCH', 'MXCC', 'MXSN', 'COBT', 
               'COCF', 'CABB', 'MXEC', 'MXDP', 'MXFR', 'MXFM', 'MXMC', 'MXPS', 'MXSO', 'MXSP', 'MXSW', 'MUST', 'NDCR', 'NOOD', 'NUGG', 'OAFL', 'OLIV', 'PANC', 'PANE', 'PAST', 'PSAU', 'PNOU', 'PORK', 'PPMX', 
               'PWSM', 'PCCE', 'DOUG', 'PPIZ', 'REFR', 'RICE', 'RBIS', 'RTEB', 'RTEM', 'SDRE', 'SALT', 'SLTC', 'SARD', 'SAUS', 'SCHN', 'SNAC', 'SNOO', 'SOUP', 'SOYS', 'SPAG', 'SPCH', 'SUGA', 'SWCO', 'SWSP', 
               'SWEE', 'TOAS', 'TOMA', 'TUNA', 'VMLK', 'WFLO', 'AIRC', 'BARS', 'BLEA', 'CBLK', 'CGLO', 'CLSP', 'CLTO', 'FILT', 'CRHC', 'CRLA', 'CRPA', 'DISH', 'DPAC', 'DRUB', 'FBRF', 'FWAX', 'FDEO', 'FRNP', 
               'GBBG', 'GCLE', 'CLEA', 'INSE', 'KITT', 'LAUN', 'LSTA', 'MXBC', 'MXHC', 'MXCB', 'MXLB', 'MXLD', 'CRTO', 'NAPK', 'PLWF', 'SCOU', 'SOFT', 'STRM', 'TOIP', 'WIPE', 'ANLG', 'FSUP', 'GMED', 'VITA', 'nan', 
               'BATT', 'CGAS', 'PFHH', 'PFIN', 'INKC', 'PETF', 'TELE', 'TILL', 'TOBA', 'ADIP', 'BSHM', 'RAZO', 'BDCR', 'CWIP', 'COMB', 'COND', 'CRHY', 'CRPC', 'DEOD', 'DIAP', 'FCCR', 'FTIS', 'FEMI', 'FRAG', 'HAIR', 
               'HRCO', 'HREM', 'HRST', 'HSTY', 'HRTR', 'LINI', 'MAKE', 'MEDS', 'CRDT', 'MXMH', 'MOWA', 'ORAL', 'SPAD', 'STOW', 'SHAM', 'SHAV', 'SKCR', 'SUNP', 'TALC', 'TAMP', 'TOIL', 'TOOB', 'TOOT', 'BAGS', 'CLPC', 
               'GRPC', 'MRKR', 'NTBK', 'SCHS', 'CSTD', 'CORP', 'CROS', 'CRBA', 'CRBR', 'CRDT', 'CRDF', 'CRFO', 'CRSA', 'CRSN', 'DEMO', 'FLSH', 'HLVW', 'COCP', 'CRSN', 'MULT', 'PCHK', 'STCK', 'MIHC','FLWT'],
})

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

('E','1'):'1W - ¿Cuándo?',
('E','3-1'):'3W - ¿Qué tamaños?',
('E','3-2'):'3W - ¿Qué marcas?',
('E','3-3'):'3W - ¿Qué sabores?',
('E','4'):'4W - ¿Quiénes? NSE (Nivel Socioeconómico)',
('E','5-1'):'5W - ¿Dónde? Regiones',
('E','5-2'):'5W - ¿Dónde? Canales',
('E','6'):'Players'
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








#Variavel de controle do numero de graficos
c_fig=0

plot=plot_ven()
last_reference_source = None
#---------------------------------------------------------------------------------------------------------------------
for w in W:

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
        left_margin = Cm(1.2)
        vertical_spacing = Cm(0.2)

        removed_headers = []

        for idx, serie in enumerate(series_configs):
            apo = aporte(serie.data.copy(), serie.pipeline, lang, serie.raw_tipo)
            removed_headers.extend(apo.attrs.get("removed_columns", []))
            c_fig += 1
            # Si hay dos tablas, ambas se posicionan a la misma altura y ocupan el espacio óptimo
            if len(series_configs) == 2:
                left_margin = Cm(1.2)
                right_margin = Cm(1.2)
                gap = Cm(0.5)
                top_position = ppt.slide_height - target_table_height - bottom_margin
                if serie.display_tipo.lower() == 'ventas':
                    left_position = ppt.slide_width / 2 + gap / 2
                else:
                    left_position = left_margin
                pic = slide.shapes.add_picture(graf_apo(apo, c_fig), left_position, top_position, height=target_table_height)
            else:
                top_position = ppt.slide_height - target_table_height - bottom_margin
                left_position = left_margin
                pic = slide.shapes.add_picture(graf_apo(apo, c_fig), left_position, top_position, height=target_table_height)
            plt.clf()

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

        if plot=="1" and len(series_configs)>1:
            df_full = series_configs[0].data.copy()
            for extra in series_configs[1:]:
                df_full = pd.concat([df_full, extra.data.iloc[:,1:]], axis=1)
            pipeline_combined = max((cfg.pipeline for cfg in series_configs), default=0)
            c_fig+=1
            left_margin = Inches(0.33)
            right_margin = Inches(0.33)
            available_line_width = available_width(ppt, left_margin, right_margin)
            pic=slide.shapes.add_picture(line_graf(df_full,pipeline_combined,titulo,c_fig,len(series_configs), width_emu=available_line_width, height_emu=Cm(10)), left_margin, Inches(1.15),width=available_line_width,height=Cm(10))
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
                pic=slide.shapes.add_picture(line_graf(serie.data,serie.pipeline,titulo+' '+serie.display_tipo,c_fig,ven_param, width_emu=chart_width, height_emu=Cm(10), multi_chart=True), left_position, Inches(1.15),width=chart_width,height=Cm(10))
                plt.clf()
        elif series_configs:
            c_fig+=1
            left_margin = Inches(0.33)
            right_margin = Inches(0.33)
            available_single_width = available_width(ppt, left_margin, right_margin)
            pic=slide.shapes.add_picture(line_graf(series_configs[0].data,series_configs[0].pipeline,titulo,c_fig,len(series_configs), width_emu=available_single_width, height_emu=Cm(10)), left_margin, Inches(1.15),width=available_single_width,height=Cm(10))
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

