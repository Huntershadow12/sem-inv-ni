# 4. SCRIPTS MSA (MARKET SHARE ANALYSIS)

**Nota:** *Los siguientes scripts están diseñados para el análisis de datos orientado a la estimación de market share, la generación de visualizaciones analíticas y la identificación de insights estratégicos a partir de datos históricos obtenidos de TradeMap. Se sugiere implementar este flujo de trabajo en entornos de ejecución Python tales como:*

- [Google Colab](https://colab.research.google.com/)
- [Python-Fiddle](https://python-fiddle.com/)
- [Kaggle Notebooks](https://www.kaggle.com/code)
- [Deepnote](https://deepnote.com/)

## 4.1. Importar e instalar librerías

```
# Instalar librerías si es necesario
!pip install openpyxl

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from google.colab import files
import warnings
warnings.filterwarnings('ignore')

# Configurar estilo de gráficos
plt.style.use('seaborn-v0_8')
sns.set_palette("husl")
```

## 4.2. Cargar archivos fuente (Trademap)

**Nota:** *El script reconoce automáticamente la columna de países y las columnas de valor por año, con cabeceras en español o inglés (p. ej. "Exportadores" / "Exporters", "Valor importado en 2024" / "Imported value in 2024"), aunque el archivo tenga filas de título encima. Acepta archivos .xlsx y .xls. Si la detección no es concluyente, el script muestra la lista numerada de columnas y pide elegirlas por número (p. ej. `0`, `1-5` o `1,3,5`) o por nombre. Para elegirlas siempre a mano, cambie `MODO_MANUAL = False` por `MODO_MANUAL = True`. La salida es la misma de siempre: `df` con las columnas `year`, `partner_country` y `trade_value_usd`. Ver [mejoras-4.2-4.5.md](mejoras-4.2-4.5.md).*

```
# =====================================================================
# 4.2. CARGAR ARCHIVOS FUENTE (TRADEMAP) — detección flexible de columnas
# Acepta cabeceras en español o inglés (Exportadores / Exporters,
# Valor importado en 2024 / Imported value in 2024, etc.).
# Si la detección automática no es concluyente, pide elegir las columnas.
# Salida: df con las columnas year, partner_country, trade_value_usd
# =====================================================================
import re

MODO_MANUAL = False   # Cambie a True para elegir siempre las columnas a mano

# Palabras clave (en minúsculas y sin tildes) para reconocer cabeceras
CLAVES_PAIS = ['exportador', 'importador', 'exporter', 'importer', 'pais', 'country',
               'countries', 'socio', 'partner', 'mercado', 'market', 'proveedor',
               'supplier', 'origen', 'origin', 'destino', 'destination', 'reporter']
CLAVES_VALOR = ['valor', 'value', 'importado', 'exportado', 'imported', 'exported',
                'importacion', 'exportacion', 'import', 'export', 'usd', 'dolar', 'dollar']
CLAVES_EXCLUIR = ['cantidad', 'quantity', 'unidad', 'unit', 'tasa', 'rate', 'growth',
                  'crecimiento', 'saldo', 'balance', 'participacion', 'share', 'precio',
                  'price', 'unitario', 'posicion', 'position', 'ranking', 'distancia',
                  'distance', 'concentracion', 'concentration', 'arancel', 'tariff', '%']
NOMBRES_MUNDO = ['mundo', 'world']
_ANIO = re.compile(r'(?<!\d)(19\d{2}|20\d{2})(?!\d)')


def _limpiar(texto):
    """Minúsculas y sin tildes, para comparar cabeceras."""
    t = str(texto).lower().strip()
    for a, b in zip('áéíóúüñ', 'aeiouun'):
        t = t.replace(a, b)
    return t


def _anio_de(cabecera):
    m = _ANIO.search(str(cabecera))
    return int(m.group(1)) if m else None


def _proporcion_numerica(serie):
    datos = serie.dropna()
    datos = datos[datos.astype(str).str.strip() != '']
    if len(datos) == 0:
        return 0.0
    return pd.to_numeric(datos, errors='coerce').notna().mean()


def _proporcion_texto(serie):
    datos = serie.dropna()
    if len(datos) == 0:
        return 0.0
    return datos.apply(lambda v: isinstance(v, str) and not str(v).strip().replace('.', '', 1).isdigit()).mean()


# ---------------------------------------------------------------------
# A) LECTURA DEL ARCHIVO (.xlsx, .xls y .xls en formato HTML de TradeMap)
# ---------------------------------------------------------------------
def leer_archivo_trademap(ruta):
    errores = []
    try:
        return pd.read_excel(ruta, sheet_name=0, header=None)
    except Exception as e:
        errores.append(f'read_excel: {e}')
    try:   # Algunas descargas .xls de TradeMap son en realidad tablas HTML
        from io import StringIO
        contenido = open(ruta, 'rb').read()
        try:
            texto = contenido.decode('utf-8')
        except UnicodeDecodeError:
            texto = contenido.decode('latin-1')
        tablas = pd.read_html(StringIO(texto))
        tabla = max(tablas, key=len)
        if not isinstance(tabla.columns, pd.RangeIndex):
            cab = [' '.join(str(x) for x in c if 'Unnamed' not in str(x)) if isinstance(c, tuple) else c
                   for c in tabla.columns]
            tabla = pd.concat([pd.DataFrame([cab], columns=tabla.columns), tabla], ignore_index=True)
        tabla.columns = range(tabla.shape[1])
        return tabla
    except Exception as e:
        errores.append(f'read_html: {e}')
    pista = ('\nSi el archivo es .xls, ejecute primero: !pip install xlrd'
             if str(ruta).lower().endswith('.xls') else '')
    raise ValueError('❌ No se pudo leer el archivo. Verifique que sea un Excel (.xlsx o .xls) '
                     'descargado de TradeMap.' + pista + '\nDetalle: ' + ' | '.join(errores))


def ubicar_cabeceras(tabla):
    """Encuentra la fila de cabeceras (TradeMap a veces pone títulos encima)."""
    mejor_fila, mejor_puntaje = 0, -1
    for i in range(min(20, len(tabla))):
        fila = tabla.iloc[i]
        textos = [str(v) for v in fila if pd.notna(v) and str(v).strip() != '']
        if len(textos) < 2:
            continue
        con_anio = sum(1 for v in textos if _anio_de(v))
        con_clave = sum(1 for v in textos if any(k in _limpiar(v) for k in CLAVES_PAIS + CLAVES_VALOR))
        puntaje = con_anio + 2 * con_clave
        if puntaje > mejor_puntaje:
            mejor_fila, mejor_puntaje = i, puntaje
    cabeceras, vistos = [], {}
    for j, v in enumerate(tabla.iloc[mejor_fila]):
        nombre = str(v).strip() if pd.notna(v) and str(v).strip() != '' else f'Columna_{j}'
        if re.fullmatch(r'\d{4}\.0', nombre):
            nombre = nombre[:-2]
        if nombre in vistos:                      # cabeceras repetidas
            vistos[nombre] += 1
            nombre = f'{nombre}_{vistos[nombre]}'
        else:
            vistos[nombre] = 0
        cabeceras.append(nombre)
    datos = tabla.iloc[mejor_fila + 1:].copy()
    datos.columns = cabeceras
    datos = datos.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)
    return datos, mejor_fila


# ---------------------------------------------------------------------
# B) DETECCIÓN AUTOMÁTICA
# ---------------------------------------------------------------------
def detectar_columnas(datos):
    """Devuelve (col_pais, columnas_valor, mensajes). Si algo no es
    concluyente, col_pais o columnas_valor quedan en None."""
    mensajes = []

    # Columnas de valor: tienen año + palabra de valor (o solo el año)
    familias = {}
    for col in datos.columns:
        anio = _anio_de(col)
        if anio is None:
            continue
        c = _limpiar(col)
        if any(k in c for k in CLAVES_EXCLUIR):
            continue
        solo_anio = re.fullmatch(r'\s*(19|20)\d{2}\s*', str(col)) is not None
        if not (solo_anio or any(k in c for k in CLAVES_VALOR)):
            continue
        if _proporcion_numerica(datos[col]) < 0.5:
            continue
        familia = _ANIO.sub('####', c)
        familias.setdefault(familia, []).append(col)

    columnas_valor = None
    if len(familias) == 1:
        columnas_valor = list(familias.values())[0]
    elif len(familias) > 1:
        mensajes.append('Se encontraron varios grupos de columnas con valores por año: '
                        + '; '.join(f'"{v[0]}"…' for v in familias.values())
                        + '. Indique cuál usar.')
    else:
        mensajes.append('No se reconocieron columnas de valor comercial por año.')

    # Columna de países: palabras clave + contenido de texto + fila Mundo/World
    candidatas = []
    for col in datos.columns:
        if columnas_valor and col in columnas_valor:
            continue
        c = _limpiar(col)
        texto = _proporcion_texto(datos[col])
        if texto < 0.5:
            continue
        valores = datos[col].dropna().astype(str).map(_limpiar)
        puntaje = 1
        if any(k in c for k in CLAVES_PAIS):
            puntaje += 2
        if valores.isin(NOMBRES_MUNDO).any():
            puntaje += 2
        if valores.nunique() >= 0.8 * len(valores):
            puntaje += 1
        candidatas.append((puntaje, col))
    candidatas.sort(reverse=True, key=lambda x: x[0])

    col_pais = None
    if not candidatas:
        mensajes.append('No se encontró una columna de texto con países.')
    elif len(candidatas) == 1 or candidatas[0][0] > candidatas[1][0]:
        if candidatas[0][0] >= 3 or len(candidatas) == 1:
            col_pais = candidatas[0][1]
        else:
            mensajes.append(f'La columna "{candidatas[0][1]}" podría ser la de países, '
                            'pero no hay certeza suficiente.')
    else:
        empatadas = [c for p, c in candidatas if p == candidatas[0][0]]
        mensajes.append('Varias columnas podrían contener los países: '
                        + ', '.join(f'"{c}"' for c in empatadas) + '.')
    return col_pais, columnas_valor, mensajes


# ---------------------------------------------------------------------
# C) VALIDACIÓN DE LAS COLUMNAS ELEGIDAS
# ---------------------------------------------------------------------
def validar_columnas(datos, col_pais, columnas_valor, anios_manuales=None):
    """Devuelve (anio_por_columna, errores, avisos)."""
    errores, avisos = [], []
    anios_manuales = anios_manuales or {}
    if col_pais not in datos.columns:
        errores.append(f'La columna de países "{col_pais}" no existe.')
    elif _proporcion_texto(datos[col_pais]) < 0.5:
        errores.append(f'La columna "{col_pais}" no parece contener nombres de países '
                       '(la mayoría de sus datos son números o están vacíos).')
    if not columnas_valor:
        errores.append('No se eligió ninguna columna de valores.')
        return {}, errores, avisos

    anio_por_columna = {}
    for col in columnas_valor:
        if col not in datos.columns:
            errores.append(f'La columna "{col}" no existe.')
            continue
        if col == col_pais:
            errores.append(f'"{col}" no puede ser a la vez columna de países y de valores.')
            continue
        anio = anios_manuales.get(col, _anio_de(col))
        if anio is None:
            errores.append(f'No se encontró el año en la cabecera "{col}".')
            continue
        p = _proporcion_numerica(datos[col])
        if p < 0.5:
            errores.append(f'La columna "{col}" no contiene valores numéricos '
                           f'(solo {p:.0%} de sus datos son números).')
            continue
        if p < 0.95:
            avisos.append(f'En "{col}" el {1 - p:.0%} de los datos no son números; '
                          'esos registros se descartarán.')
        anio_por_columna[col] = anio

    anios = list(anio_por_columna.values())
    repetidos = sorted({a for a in anios if anios.count(a) > 1})
    if repetidos:
        errores.append(f'Hay más de una columna para el/los año(s) {repetidos}. '
                       'Elija una sola columna por año.')
    return anio_por_columna, errores, avisos


# ---------------------------------------------------------------------
# D) ASIGNACIÓN MANUAL (por número o por nombre de columna)
# ---------------------------------------------------------------------
def _mostrar_columnas(datos):
    print('\n📋 Columnas disponibles en el archivo:')
    print(f'  {"N.º":>4}  {"Cabecera":<45} Ejemplo')
    for i, col in enumerate(datos.columns):
        ejemplo = datos[col].dropna()
        ejemplo = str(ejemplo.iloc[0])[:25] if len(ejemplo) else '(vacía)'
        print(f'  {i:>4}  {str(col)[:45]:<45} {ejemplo}')


def _interpretar(entrada, datos):
    """Convierte '0', 'Exporters', '2-6' o '2,3,5' en nombres de columnas."""
    cols = list(datos.columns)
    resultado = []
    for parte in [p.strip() for p in entrada.split(',') if p.strip()]:
        if re.fullmatch(r'\d+\s*-\s*\d+', parte):
            a, b = [int(x) for x in parte.split('-')]
            if not (0 <= a <= b < len(cols)):
                raise ValueError(f'El rango "{parte}" está fuera de 0–{len(cols) - 1}.')
            resultado += cols[a:b + 1]
        elif parte.isdigit():
            if int(parte) >= len(cols):
                raise ValueError(f'No existe la columna número {parte}.')
            resultado.append(cols[int(parte)])
        elif parte in cols:
            resultado.append(parte)
        else:
            raise ValueError(f'No existe una columna llamada "{parte}".')
    return resultado


def asignacion_manual(datos, sugerida_pais=None, sugeridas_valor=None):
    _mostrar_columnas(datos)
    while True:
        try:
            print('\n✍️  Asignación manual (escriba el número o el nombre de la columna):')
            txt = input(f'  Columna de PAÍSES / socios'
                        f'{f" [Enter = {sugerida_pais}]" if sugerida_pais else ""}: ').strip()
            col_pais = sugerida_pais if (not txt and sugerida_pais) else _interpretar(txt, datos)[0]

            ayuda = f' [Enter = {len(sugeridas_valor)} sugeridas]' if sugeridas_valor else ''
            txt = input(f'  Columnas de VALOR por año (ej. 1-5 o 1,2,3){ayuda}: ').strip()
            columnas_valor = sugeridas_valor if (not txt and sugeridas_valor) else _interpretar(txt, datos)

            anios_manuales = {}
            for col in columnas_valor:
                if _anio_de(col) is None and col != col_pais:
                    a = input(f'  La cabecera "{col}" no indica el año. Escriba el año: ').strip()
                    if not re.fullmatch(r'(19|20)\d{2}', a):
                        raise ValueError(f'"{a}" no es un año válido.')
                    anios_manuales[col] = int(a)

            anio_por_columna, errores, avisos = validar_columnas(datos, col_pais, columnas_valor, anios_manuales)
            if errores:
                for e in errores:
                    print('  ❌ ' + e)
                print('  Intente de nuevo.')
                continue
            return col_pais, anio_por_columna, avisos
        except (ValueError, IndexError) as e:
            print(f'  ❌ {e} Intente de nuevo.')


# ---------------------------------------------------------------------
# E) EJECUCIÓN
# ---------------------------------------------------------------------
# Subir archivo Excel directamente desde TradeMap
try:
    uploaded = files.upload()
    filename = list(uploaded.keys())[0]
except NameError:          # fuera de Google Colab
    filename = input('Ruta del archivo de TradeMap: ').strip()

# Leer datos tal como vienen de TradeMap
tabla_original = leer_archivo_trademap(filename)
df_raw, fila_cabecera = ubicar_cabeceras(tabla_original)

# Verificar estructura original
if fila_cabecera > 0:
    print(f'ℹ️  Se omitieron {fila_cabecera} fila(s) de título antes de las cabeceras.')
print("Columnas disponibles:", df_raw.columns.tolist())
print("\nPrimeras 5 filas:")
print(df_raw.head())

# Identificar columna de países y columnas de valor por año
col_pais, columnas_valor, mensajes = detectar_columnas(df_raw)
anio_por_columna, errores, avisos = ({}, ['Detección no concluyente.'], [])
if col_pais and columnas_valor and not MODO_MANUAL:
    anio_por_columna, errores, avisos = validar_columnas(df_raw, col_pais, columnas_valor)

if errores or MODO_MANUAL:
    for m in mensajes + [e for e in errores if e != 'Detección no concluyente.']:
        print('⚠️  ' + m)
    print('\n🔎 La detección automática no fue concluyente: elija las columnas manualmente.'
          if not MODO_MANUAL else '\n✍️  Modo manual activado.')
    col_pais, anio_por_columna, avisos = asignacion_manual(df_raw, col_pais, columnas_valor)
else:
    print(f'\n✅ Detección automática: países = "{col_pais}"')
for a in avisos:
    print('⚠️  ' + a)

year_columns = list(anio_por_columna)
print(f"\nColumnas de años encontradas: {year_columns}")

# Transformar de formato ancho a formato largo (melt)
df_melted = df_raw.melt(
    id_vars=[col_pais],
    value_vars=year_columns,
    var_name='year_column',
    value_name='trade_value_usd'
)

# Asignar el año de cada columna
df_melted['year'] = df_melted['year_column'].map(anio_por_columna).astype(int)
df_melted = df_melted.rename(columns={col_pais: 'partner_country'})

# Limpiar datos
df_melted['trade_value_usd'] = pd.to_numeric(df_melted['trade_value_usd'], errors='coerce')
df = df_melted[['year', 'partner_country', 'trade_value_usd']].copy()

# Eliminar filas con datos faltantes y fila "Mundo" / "World"
df = df.dropna(subset=['trade_value_usd', 'partner_country'])
df['partner_country'] = df['partner_country'].astype(str).str.strip()
df = df[~df['partner_country'].map(_limpiar).isin(NOMBRES_MUNDO)].reset_index(drop=True)

if df.empty:
    raise ValueError('❌ No quedaron datos después de la limpieza. Revise las columnas elegidas.')

print(f"\nDatos procesados: {len(df)} filas")
print(f"Países únicos: {df['partner_country'].nunique()}")
print(f"Años disponibles: {[int(a) for a in sorted(df['year'].unique())]}")
print("\nPrimeros datos procesados:")
print(df.head(10))
```

## 4.3. Calcular Market Share, Ranking y aplicar filtros

```
# Calcular market share por año y país
market_evolution = df.groupby(['year', 'partner_country'])['trade_value_usd'].sum().reset_index()

# Calcular market share total por año
yearly_totals = market_evolution.groupby('year')['trade_value_usd'].sum()
market_evolution['market_share'] = market_evolution.apply(
    lambda x: (x['trade_value_usd'] / yearly_totals[x['year']]) * 100, axis=1
)

# Identificar top países
top_countries = df.groupby('partner_country')['trade_value_usd'].sum().nlargest(8).index

# Filtrar solo top países
market_evolution_top = market_evolution[market_evolution['partner_country'].isin(top_countries)]
```

## 4.4.  Generar Tablas Pivot para análisis

```
# TABLA 1: Market Share por año y país
print("📊 TABLA 1: MARKET SHARE POR AÑO (%)")
print("="*60)
pivot_share = market_evolution_top.pivot(index='year', columns='partner_country', values='market_share')
pivot_share_rounded = pivot_share.round(1)
print(pivot_share_rounded.to_string())

# TABLA 2: Valores de comercio (Millones USD)
print("\n\n💰 TABLA 2: VALORES DE COMERCIO (Millones USD)")
print("="*60)
pivot_value = market_evolution_top.pivot(index='year', columns='partner_country', values='trade_value_usd')
pivot_value_millions = (pivot_value / 1000).round(1)  # Convertir a millones
print(pivot_value_millions.to_string())

# TABLA 3: Crecimiento anual por país
print("\n\n📈 TABLA 3: CRECIMIENTO ANUAL COMPUESTO (CAGR %)")
print("="*60)
growth_data = market_evolution_top.groupby('partner_country').apply(
    lambda x: ((x['trade_value_usd'].iloc[-1] / x['trade_value_usd'].iloc[0]) ** (1/(len(x)-1)) - 1) * 100
).reset_index(name='cagr')
growth_summary = growth_data.set_index('partner_country')['cagr'].round(1).sort_values(ascending=False)
print(growth_summary.to_string())

# TABLA 4: Ranking por año
print("\n\n🏆 TABLA 4: RANKING DE PAÍSES POR AÑO (Top 5)")
print("="*60)
for year in sorted(df['year'].unique()):
    year_data = market_evolution_top[market_evolution_top['year'] == year].nlargest(5, 'market_share')
    print(f"\n{year}:")
    for i, (_, row) in enumerate(year_data.iterrows(), 1):
        print(f"  {i}. {row['partner_country']}: {row['market_share']:.1f}% ({row['trade_value_usd']/1000:.0f}M USD)")
```

## 4.5. Generar Gráficos analíticos

**Nota:** *Al ejecutar esta celda aparece un selector de paleta de colores con tres opciones predeterminadas (Ejecutivo Clásico, Comercio Global y Corporativo Contemporáneo) y una opción personalizada para pegar código CSS, por ejemplo el que exporta [Coolors](https://coolors.co/) (Export → CSS). Elija una opción y pulse **Generar gráficos**; para cambiar de paleta, elija otra y vuelva a pulsar el botón, sin modificar el código. El CSS no se ejecuta: solo se extraen y validan los códigos HEX. Ver [mejoras-4.2-4.5.md](mejoras-4.2-4.5.md).*

```
# =====================================================================
# 4.5. GENERAR GRÁFICOS ANALÍTICOS — con selector de paleta de colores
# Requiere haber ejecutado antes las secciones 4.1 a 4.4
# (usa: df, market_evolution_top, top_countries, pivot_share,
#  pivot_value y growth_data).
# =====================================================================
import re
import matplotlib.colors as mcolors
from IPython.display import display, HTML

# ---------------------------------------------------------------------
# A) DEFINICIÓN CENTRALIZADA DE PALETAS
#    Para cambiar o añadir una paleta, edite solo este diccionario.
#    'series' : colores asignados a los países (líneas, marcadores, barras)
#    'texto'  : color de etiquetas, cuadrícula y línea de referencia
# ---------------------------------------------------------------------
PALETAS = {
    'Ejecutivo Clásico': {
        'descripcion': 'Azules, grises y acentos sobrios para informes ejecutivos.',
        'series': ['#1F3A5F', '#3D6A99', '#7FA7CF', '#4A4A4A',
                   '#9AA5B1', '#C9A227', '#B5543C', '#2E8B8B'],
        'texto': '#2B2B2B',
    },
    'Comercio Global': {
        'descripcion': 'Azules y verdes que evocan negocios internacionales.',
        'series': ['#0B3C5D', '#1D70A2', '#2BB3C0', '#0F8B6E',
                   '#6BBF59', '#A7D7C5', '#F2A541', '#5C6B73'],
        'texto': '#1C2B36',
    },
    'Corporativo Contemporáneo': {
        'descripcion': 'Neutros, tierras y acentos cálidos para reportes modernos.',
        'series': ['#3E3A36', '#8C7B6B', '#B8A99A', '#A0522D',
                   '#C97C5D', '#D9A441', '#6B7F5E', '#E07A5F'],
        'texto': '#2F2A26',
    },
}
OPCION_PERSONALIZADA = 'Personalizada (pegar código CSS)'
MIN_COLORES_PERSONALIZADOS = 2


# ---------------------------------------------------------------------
# B) LECTURA SEGURA DE CÓDIGO CSS
#    El CSS NUNCA se ejecuta: solo se buscan códigos HEX con
#    expresiones regulares y se validan con Matplotlib.
# ---------------------------------------------------------------------
_HEX_VALIDO = re.compile(r'#([0-9a-fA-F]{8}|[0-9a-fA-F]{6}|[0-9a-fA-F]{3,4})(?![0-9a-zA-Z_-])')
_TOKEN_HEX = re.compile(r'#[0-9A-Za-z]+')
_VARIABLE_CSS = re.compile(r'--[\w-]+\s*:\s*(#[0-9A-Za-z]+)')


def _normalizar_hex(codigo):
    """Convierte #abc / #abcd / #aabbcc / #aabbccdd en #AABBCC.
    Devuelve (hex_6_digitos, tenia_transparencia)."""
    h = codigo.lstrip('#')
    if len(h) in (3, 4):
        h = ''.join(c * 2 for c in h)
    transparencia = len(h) == 8 and h[6:].lower() != 'ff'
    return '#' + h[:6].upper(), transparencia


def extraer_colores_css(texto_css):
    """Extrae los colores HEX de un bloque CSS (p. ej. exportado de Coolors).
    Devuelve (colores, avisos, errores)."""
    avisos, errores = [], []
    if not texto_css or not texto_css.strip():
        return [], avisos, ['No se pegó ningún código CSS.']

    sin_comentarios = re.sub(r'/\*.*?\*/', ' ', texto_css, flags=re.S)

    # 1) Preferir variables CSS (--nombre: #hex), como en :root de Coolors
    candidatos = _VARIABLE_CSS.findall(sin_comentarios)
    # 2) Si no hay variables, tomar cualquier código HEX del texto
    if not candidatos:
        candidatos = _TOKEN_HEX.findall(sin_comentarios)

    colores, invalidos, con_alfa = [], [], []
    for token in candidatos:
        if not _HEX_VALIDO.fullmatch(token):
            invalidos.append(token)
            continue
        hex6, alfa = _normalizar_hex(token)
        if not mcolors.is_color_like(hex6):
            invalidos.append(token)
            continue
        if alfa:
            con_alfa.append(token)
        if hex6 not in colores:          # quitar repetidos, conservar el orden
            colores.append(hex6)

    if invalidos:
        avisos.append('Se ignoraron códigos HEX no válidos: ' + ', '.join(invalidos))
    if con_alfa:
        avisos.append('Se ignoró la transparencia de: ' + ', '.join(con_alfa)
                      + ' (los gráficos usan colores sólidos).')
    if not colores:
        errores.append('El código no contiene colores HEX válidos '
                       '(formato esperado: #RRGGBB, p. ej. --azul: #1D70A2;).')
    elif len(colores) < MIN_COLORES_PERSONALIZADOS:
        errores.append(f'Solo se encontró {len(colores)} color válido; '
                       f'se necesitan al menos {MIN_COLORES_PERSONALIZADOS}.')
    return colores, avisos, errores


def _luminancia(hex_color):
    r, g, b = mcolors.to_rgb(hex_color)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _variar(hex_color, factor):
    """factor > 0 aclara (mezcla con blanco); factor < 0 oscurece."""
    rgb = mcolors.to_rgb(hex_color)
    if factor >= 0:
        rgb = [c + (1 - c) * factor for c in rgb]
    else:
        rgb = [c * (1 + factor) for c in rgb]
    return mcolors.to_hex(rgb).upper()


def construir_paleta_personalizada(colores, n_necesarios):
    """Arma la paleta a partir de los colores del usuario.
    Si hay menos colores que países, completa con variaciones (más claras
    u oscuras) de esos mismos colores e informa al usuario."""
    avisos = []
    series = list(colores[:n_necesarios])
    if len(colores) < n_necesarios:
        faltan = n_necesarios - len(colores)
        intensidades = [0.40, 0.60, 0.25, 0.75]
        k = 0
        while len(series) < n_necesarios:
            base = colores[k % len(colores)]
            f = intensidades[(k // len(colores)) % len(intensidades)]
            # Colores oscuros se aclaran; colores claros se oscurecen
            nuevo = _variar(base, f if _luminancia(base) < 0.55 else -f)
            if nuevo not in series:
                series.append(nuevo)
            k += 1
            if k > 200:
                break
        avisos.append(f'La paleta tiene {len(colores)} colores y hay {n_necesarios} países: '
                      f'se generaron {faltan} variaciones de sus mismos colores.')
    elif len(colores) > n_necesarios:
        avisos.append(f'Se usaron los primeros {n_necesarios} colores de {len(colores)} '
                      f'(uno por país).')

    muy_claros = [c for c in series if _luminancia(c) > 0.85]
    if muy_claros:
        avisos.append('Colores muy claros que pueden verse poco sobre el fondo: '
                      + ', '.join(muy_claros) + '. Las barras llevarán un contorno para distinguirlas.')

    texto = min(colores, key=_luminancia)
    if _luminancia(texto) > 0.5:
        avisos.append(f'Todos los colores son claros; para que los textos se lean '
                      f'se usará {PALETAS["Ejecutivo Clásico"]["texto"]} en etiquetas.')
        texto = PALETAS['Ejecutivo Clásico']['texto']
    return {'descripcion': 'Paleta suministrada por el usuario.',
            'series': series, 'texto': texto}, avisos


def muestras_html(colores):
    chips = ''.join(f'<span title="{c}" style="display:inline-block;width:22px;height:22px;'
                    f'margin-right:3px;border-radius:4px;background:{c};'
                    f'border:1px solid #0002"></span>' for c in colores)
    return f'<div style="margin:2px 0 8px 0">{chips}</div>'


# ---------------------------------------------------------------------
# C) GENERACIÓN DE LOS 4 GRÁFICOS (mismos cálculos que la versión original;
#    solo cambian los colores, que vienen de la paleta elegida)
# ---------------------------------------------------------------------
def generar_graficos(paleta):
    global fig, axes, latest_year, latest_data, bars, bars2, colores_pais

    texto = paleta['texto']
    # Un color fijo por país, el mismo en los cuatro gráficos
    colores_pais = {pais: paleta['series'][i % len(paleta['series'])]
                    for i, pais in enumerate(top_countries)}

    def colores_de(paises):
        return [colores_pais.get(p, texto) for p in paises]

    def estilo(ax, titulo, ylabel):
        ax.set_title(titulo, fontsize=14, fontweight='bold', color=texto)
        ax.set_ylabel(ylabel, color=texto)
        ax.tick_params(colors=texto)
        ax.xaxis.label.set_color(texto)
        for borde in ax.spines.values():
            borde.set_color(texto)
            borde.set_alpha(0.3)

    fig, axes = plt.subplots(2, 2, figsize=(16, 14))

    # 1. Market Share Trends con valores en puntos clave
    pivot_share.plot(kind='line', marker='o', ax=axes[0,0], linewidth=2, markersize=6,
                     color=colores_de(pivot_share.columns))
    estilo(axes[0,0], 'Market Share Evolution (%)', 'Market Share (%)')
    axes[0,0].grid(True, alpha=0.3, color=texto)
    axes[0,0].legend(bbox_to_anchor=(1.05, 1), loc='upper left', labelcolor=texto)

    # Añadir valores para Brasil y Colombia (principales)
    for country in ['Brasil', 'Colombia']:
        if country in pivot_share.columns:
            for year in pivot_share.index:
                value = pivot_share.loc[year, country]
                if not pd.isna(value):
                    axes[0,0].annotate(f'{value:.1f}%', (year, value),
                                       textcoords="offset points", xytext=(0,10),
                                       ha='center', fontsize=8, color=texto)

    # 2. Trade Value Evolution
    (pivot_value/1000).plot(kind='line', marker='s', ax=axes[0,1], linewidth=2, markersize=6,
                            color=colores_de(pivot_value.columns))
    estilo(axes[0,1], 'Trade Value Evolution (Million USD)', 'Million USD')
    axes[0,1].grid(True, alpha=0.3, color=texto)
    axes[0,1].legend(labelcolor=texto)

    # 3. Market Share actual con valores
    latest_year = df['year'].max()
    latest_data = market_evolution_top[market_evolution_top['year'] == latest_year].nlargest(8, 'market_share')
    serie_actual = latest_data.set_index('partner_country')['market_share']
    bars = serie_actual.plot(kind='bar', ax=axes[1,0], color=colores_de(serie_actual.index),
                            edgecolor=texto, linewidth=0.6)
    estilo(axes[1,0], f'Market Share {latest_year} (%)', 'Market Share (%)')
    axes[1,0].tick_params(axis='x', rotation=45)
    for i, (country, value) in enumerate(serie_actual.items()):
        axes[1,0].text(i, value + 0.5, f'{value:.1f}%', ha='center', va='bottom',
                       fontsize=9, color=texto)

    # 4. CAGR con valores
    if 'growth_data' in globals():
        serie_cagr = growth_data.set_index('partner_country')['cagr']
        bars2 = serie_cagr.plot(kind='bar', ax=axes[1,1], color=colores_de(serie_cagr.index),
                                edgecolor=texto, linewidth=0.6)
        axes[1,1].axhline(y=0, color=texto, linestyle='-', alpha=0.4)
        for i, (country, value) in enumerate(serie_cagr.items()):
            if pd.isna(value):
                continue
            axes[1,1].text(i, value + 0.5 if value > 0 else value - 1, f'{value:.1f}%',
                           ha='center', va='bottom' if value > 0 else 'top',
                           fontsize=9, color=texto)
    else:
        bars2 = axes[1,1]
        axes[1,1].text(0.5, 0.5, 'CAGR no disponible\n(se requieren al menos dos años de datos)',
                       ha='center', va='center', transform=axes[1,1].transAxes, color=texto)
    estilo(axes[1,1], 'CAGR % (Compound Annual Growth Rate)', 'CAGR (%)')
    axes[1,1].tick_params(axis='x', rotation=45)

    plt.tight_layout()
    plt.show()


# ---------------------------------------------------------------------
# D) INTERFAZ INTERACTIVA (Google Colab / Jupyter)
# ---------------------------------------------------------------------
def resolver_paleta(opcion, texto_css=''):
    """Devuelve (paleta, avisos, errores) según la opción elegida."""
    if opcion in PALETAS:
        return PALETAS[opcion], [], []
    colores, avisos, errores = extraer_colores_css(texto_css)
    if errores:
        return None, avisos, errores
    paleta, avisos_extra = construir_paleta_personalizada(colores, len(top_countries))
    return paleta, avisos + avisos_extra, []


def ejecutar_con_paleta(opcion, texto_css=''):
    global paleta_seleccionada
    paleta, avisos, errores = resolver_paleta(opcion, texto_css)
    for a in avisos:
        print('⚠️  ' + a)
    if errores:
        for e in errores:
            print('❌ ' + e)
        print('Corrija el código CSS y vuelva a intentarlo. No se generaron gráficos.')
        return
    paleta_seleccionada = {'nombre': opcion, **paleta}
    print(f'✅ Paleta aplicada: {opcion}')
    display(HTML(muestras_html(paleta['series'])))
    generar_graficos(paleta)


try:
    import ipywidgets as widgets

    opciones = list(PALETAS) + [OPCION_PERSONALIZADA]
    selector = widgets.RadioButtons(options=opciones, value=opciones[0],
                                    description='Paleta:', layout=widgets.Layout(width='auto'))
    vista_previa = widgets.HTML()
    campo_css = widgets.Textarea(
        placeholder=':root {\n  --color-1: #264653;\n  --color-2: #2a9d8f;\n  --color-3: #e9c46a;\n}',
        description='CSS:', layout=widgets.Layout(width='600px', height='160px', display='none'))
    boton = widgets.Button(description='Generar gráficos', button_style='primary', icon='bar-chart')
    salida = widgets.Output()

    def _actualizar_vista(_=None):
        if selector.value in PALETAS:
            p = PALETAS[selector.value]
            vista_previa.value = f'<i>{p["descripcion"]}</i>' + muestras_html(p['series'])
            campo_css.layout.display = 'none'
        else:
            vista_previa.value = ('<i>Pegue el código CSS de su paleta (por ejemplo, '
                                  'exportado desde Coolors: Export → CSS).</i>')
            campo_css.layout.display = 'flex'

    def _al_hacer_clic(_):
        salida.clear_output()
        with salida:
            ejecutar_con_paleta(selector.value, campo_css.value)

    selector.observe(_actualizar_vista, names='value')
    boton.on_click(_al_hacer_clic)
    _actualizar_vista()
    display(widgets.VBox([widgets.HTML('<b>🎨 Seleccione la paleta de colores para los gráficos</b>'),
                          selector, vista_previa, campo_css, boton, salida]))

except ImportError:
    # Alternativa sin widgets: menú de texto
    opciones = list(PALETAS) + [OPCION_PERSONALIZADA]
    print('🎨 Seleccione la paleta de colores:')
    for i, nombre in enumerate(opciones, 1):
        print(f'  {i}. {nombre}')
    eleccion = input('Número de opción: ').strip()
    while eleccion not in [str(i) for i in range(1, len(opciones) + 1)]:
        eleccion = input(f'Opción no válida. Escriba un número del 1 al {len(opciones)}: ').strip()
    opcion = opciones[int(eleccion) - 1]
    css = ''
    if opcion == OPCION_PERSONALIZADA:
        print('Pegue el código CSS y escriba FIN en una línea aparte para terminar:')
        lineas = []
        while (linea := input()).strip().upper() != 'FIN':
            lineas.append(linea)
        css = '\n'.join(lineas)
    ejecutar_con_paleta(opcion, css)
```

## 4.6. Generar Insights estratégicos o "Descubrimientos relevantes" 

**Nota:** *Reemplazar "producto analizado" (Line 2) y "Producto" (Line 35) por el número de partida y nombre del producto analizado antes de proceder con la ejecución del script.*

```
# Generar insights automáticos
def generate_insights(df, market_evolution_top, product_name="producto analizado"):
    latest_year = df['year'].max()
    first_year = df['year'].min()

    print(f"📊 INSIGHTS ESTRATÉGICOS - {product_name.upper()}")
    print("="*60)

    # Top 3 países
    top3 = market_evolution_top[market_evolution_top['year'] == latest_year].nlargest(3, 'market_share')
    print(f"\n🏆 TOP 3 PAÍSES ({latest_year}):")
    for i, (_, row) in enumerate(top3.iterrows(), 1):
        print(f"  {i}. {row['partner_country']}: {row['market_share']:.1f}% market share")

    # Mayor crecimiento
    growth_analysis = market_evolution_top.groupby('partner_country').apply(
        lambda x: x['trade_value_usd'].iloc[-1] / x['trade_value_usd'].iloc[0] - 1
    ).sort_values(ascending=False)

    print(f"\n📈 MAYOR CRECIMIENTO ({first_year}-{latest_year}):")
    print(f"  {growth_analysis.index[0]}: +{growth_analysis.iloc[0]:.1%}")

    # Oportunidades
    declining = growth_analysis[growth_analysis < 0]
    if len(declining) > 0:
        print(f"\n⚠️  MERCADOS EN DECLIVE:")
        for country, decline in declining.head(3).items():
            print(f"  {country}: {decline:.1%}")

    # Tamaño total del mercado
    total_market = df[df['year'] == latest_year]['trade_value_usd'].sum()
    print(f"\n💰 TAMAÑO TOTAL DEL MERCADO ({latest_year}): ${total_market/1e9:.1f}B USD")

# Define the product_name variable
product_name = "Producto"

# Aplicar análisis con nombre de producto dinámico
generate_insights(df, market_evolution_top, product_name)
```

## 4.7. Generación de Promts para análisis posteriores

**Nota:** *Reemplazar "market_name" (Line 7) por el mercado objetivo (país-destino) antes de proceder con la ejecución del script.*

```
# GENERAR PROMPT AUTOMÁTICO PARA ANÁLISIS
def generate_analysis_prompt(df, market_evolution_top, product_name):
    latest_year = df['year'].max()
    first_year = df['year'].min()

    # Detectar mercado objetivo (desde columnas o contexto)
    market_name = "Estados Unidos"  # Por defecto, se podría dinamizar

    # Generar estadísticas clave
    top3 = market_evolution_top[market_evolution_top['year'] == latest_year].nlargest(3, 'market_share')
    total_market = df[df['year'] == latest_year]['trade_value_usd'].sum()

    prompt = f"""
    Actúa como consultor en inteligencia de mercados y analiza estos datos de market share de {product_name.lower()} en {market_name} ({first_year}-{latest_year}). Proporciona insights estratégicos específicos:

    CONTEXTO DEL MERCADO:
    - Producto: {product_name}
    - Mercado objetivo: {market_name}
    - Período: {first_year}-{latest_year}
    - Tamaño total del mercado ({latest_year}): ${total_market/1000:.0f} millones USD

    TOP 3 LÍDERES ACTUALES ({latest_year}):
    """

    for i, (_, row) in enumerate(top3.iterrows(), 1):
        prompt += f"\n    {i}. {row['partner_country']}: {row['market_share']:.1f}% market share"

    prompt += f"""

    DATOS PARA ANÁLISIS:
    [Pegar aquí las 4 tablas generadas por el script anterior]

    PREGUNTAS ESPECÍFICAS A RESPONDER:

    1. ANÁLISIS COMPETITIVO:
       - ¿Cuáles son las tendencias de los líderes del mercado?
       - ¿Qué países están ganando/perdiendo participación?
       - ¿Hay oportunidades por concentración/fragmentación del mercado?

    2. OPORTUNIDADES DE ENTRADA:
       - ¿Qué nichos o segmentos presentan menos competencia?
       - ¿Existen países que han reducido su participación creando vacíos?
       - ¿Cuál sería el mejor timing para entrar al mercado?

    3. BENCHMARKING REGIONAL:
       - ¿Cómo se comportan países con perfil similar al de tu empresa?
       - ¿Qué países de América Latina tienen mejor performance?
       - ¿Cuáles son las mejores prácticas observables?

    4. ESTRATEGIA DE POSICIONAMIENTO:
       - ¿Competir directamente con líderes o buscar nichos?
       - ¿Qué ventajas competitivas serían más relevantes?
       - ¿Cuál sería un market share objetivo realista en 2-3 años?

    5. RIESGOS Y AMENAZAS:
       - ¿Qué competidores emergentes representan amenazas?
       - ¿Hay signos de saturación o madurez del mercado?
       - ¿Qué factores externos podrían afectar la industria?

    FORMATO DE RESPUESTA:
    - Insights concisos y accionables
    - Recomendaciones estratégicas específicas
    - Identificación de 2-3 oportunidades prioritarias
    - Métricas clave a monitorear
    """

    return prompt

# Generar y mostrar el prompt
analysis_prompt = generate_analysis_prompt(df, market_evolution_top, product_name)
print("="*80)
print("📋 PROMPT PARA ANÁLISIS ESTRATÉGICO:")
print("="*80)
print(analysis_prompt)
print("\n" + "="*80)
print("💡 INSTRUCCIONES:")
print("1. Copia este prompt")
print("2. Pega las 4 tablas de datos generadas arriba")
print("3. Úsalo con cualquier chatbot para obtener insights estratégicos")
print("="*80)
```

## 4.8. Helpers (opcionales) para depuración y personalización
### 4.8.1. Promt para análisis de oportunidad: Origen (Proveedor)-Destino

```
Con base en lo anterior, redacta un texto argumentativo, en tono formal y académico en tercera persona, máximo una cuartilla (250 palabras) donde evidencies los resultados de este estudio de mercados. Enfatiza en el nivel de oportunidad que tendría [COLOMBIA] para ingresar a este mercado desde el punto de vista comercial basándose únicamente en los datos suministrados.
```

### 4.8.2. Promt para personalizar colores en los gráficos

```
Con base en la siguiente paleta de colores:

[CSS.CODE]

Ajusta el siguiente trozo de código en consecuencia, da por hecho que las librerías ya se encuentran instaladas:

[PY.CODE.4.5]
```

### 4.8.3. Ajuste para datos faltantes o sin datos históricos (No calcula CAGR)

```
# TABLA 1: Market Share por año y país
print("📊 TABLA 1: MARKET SHARE POR AÑO (%)")
print("="*60)
pivot_share = market_evolution_top.pivot(index='year', columns='partner_country', values='market_share')
pivot_share_rounded = pivot_share.round(1)
print(pivot_share_rounded.to_string())

# TABLA 2: Valores de comercio (Millones USD)
print("\n\n💰 TABLA 2: VALORES DE COMERCIO (Millones USD)")
print("="*60)
pivot_value = market_evolution_top.pivot(index='year', columns='partner_country', values='trade_value_usd')
pivot_value_millions = (pivot_value / 1000).round(1)  # Convertir a millones
print(pivot_value_millions.to_string())

# TABLA 3: Crecimiento anual por país
# El cálculo del CAGR requiere al menos dos años de datos. Como solo se tiene un año (2024),
# esta tabla no se puede generar con los datos actuales.
# print("\n\n📈 TABLA 3: CRECIMIENTO ANUAL COMPUESTO (CAGR %)")
# print("="*60)
# growth_data = market_evolution_top.groupby('partner_country').apply(
#     lambda x: ((x['trade_value_usd']. iloc[-1] / x['trade_value_usd'].iloc[0]) ** (1/(len(x)-1)) - 1) * 100
# ).reset_index(name='cagr')
# growth_summary = growth_data.set_index('partner_country')['cagr'].round(1).sort_values(ascending=False)
# print(growth_summary.to_string())
print("\n\n📈 TABLA 3: CRECIMIENTO ANUAL COMPUESTO (CAGR %)")
print("="*60)
print("No se puede calcular el CAGR. Se requiere al menos dos años de datos para este cálculo.")

# TABLA 4: Ranking por año
print("\n\n🏆 TABLA 4: RANKING DE PAÍSES POR AÑO (Top 5)")
print("="*60)
for year in sorted(df['year'].unique()):
    year_data = market_evolution_top[market_evolution_top['year'] == year].nlargest(5, 'market_share')
    print(f"\n{year}:")
    for i, (_, row) in enumerate(year_data.iterrows(), 1):
        print(f"  {i}. {row['partner_country']}: {row['market_share']:.1f}% ({row['trade_value_usd']/1000:.0f}M USD)")
```

### 4.8.4. Ajuste para datos faltantes (Calcula CAGR omitiendo datos vacíos o nulos)

```
# TABLA 1: Market Share por año y país
print("📊 TABLA 1: MARKET SHARE POR AÑO (%)")
print("="*60)
pivot_share = market_evolution_top.pivot(index='year', columns='partner_country', values='market_share')
pivot_share_rounded = pivot_share.round(1)
print(pivot_share_rounded.to_string())

# TABLA 2: Valores de comercio (Millones USD)
print("\n\n💰 TABLA 2: VALORES DE COMERCIO (Millones USD)")
print("="*60)
pivot_value = market_evolution_top.pivot(index='year', columns='partner_country', values='trade_value_usd')
pivot_value_millions = (pivot_value / 1000).round(1)  # Convertir a millones
print(pivot_value_millions.to_string())

# TABLA 3: Crecimiento anual por país
print("\n\n📈 TABLA 3: CRECIMIENTO ANUAL COMPUESTO (CAGR %)")
print("="*60)
growth_data = market_evolution_top.groupby('partner_country').apply(
    lambda x: ((x['trade_value_usd'].iloc[-1] / x['trade_value_usd'].iloc[0]) ** (1/(len(x)-1)) - 1) * 100 if len(x) > 1 else np.nan
).reset_index(name='cagr')
growth_summary = growth_data.set_index('partner_country')['cagr'].round(1).sort_values(ascending=False)
print(growth_summary.to_string())

# TABLA 4: Ranking por año
print("\n\n🏆 TABLA 4: RANKING DE PAÍSES POR AÑO (Top 5)")
print("="*60)
for year in sorted(df['year'].unique()):
    year_data = market_evolution_top[market_evolution_top['year'] == year].nlargest(5, 'market_share')
    print(f"\n{year}:")
    for i, (_, row) in enumerate(year_data.iterrows(), 1):
        print(f"  {i}. {row['partner_country']}: {row['market_share']:.1f}% ({row['trade_value_usd']/1000:.0f}M USD)")
```
