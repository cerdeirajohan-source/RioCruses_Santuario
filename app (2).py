"""
Geovisualizador del Santuario de la Naturaleza Río Cruces (Valdivia, Chile)
Streamlit + Folium

Estructura esperada del repositorio:
    app.py
    requirements.txt
    data/
        Fuentes_y_proyectos_Santuario_Rio_Cruces.xlsx
        area_influencia_conaf2019_aprox.geojson
        santuario_rio_cruces.geojson      <- (opcional) límite real del santuario
"""

import html
import io
import json
import re
from pathlib import Path

import folium
import pandas as pd
import streamlit as st
from folium.plugins import Fullscreen, MarkerCluster
from streamlit_folium import st_folium

# ──────────────────────────────────────────────────────────────────────────────
# 1. CONFIGURACIÓN
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Santuario Río Cruces · Geovisualizador",
    page_icon="🦢",
    layout="wide",
)

BASE = Path(__file__).parent
ARCHIVO_EXCEL = "Fuentes_y_proyectos_Santuario_Rio_Cruces.xlsx"
ARCHIVO_INFLUENCIA = "area_influencia_conaf2019_aprox.geojson"
ARCHIVOS_SANTUARIO = ["santuario_rio_cruces.geojson", "santuario.geojson", "limite_santuario.geojson"]
HOJA_PROYECTOS = "Proyectos influencia"

CENTRO_SANTUARIO = [-39.69, -73.17]  # Valdivia, Santuario Río Cruces
ZOOM_INICIAL = 10

COLOR_SANTUARIO = "#87CEFA"   # celeste claro
COLOR_INFLUENCIA = "#4B0F7A"  # morado oscuro

# ──────────────────────────────────────────────────────────────────────────────
# GEOMETRÍAS INCRUSTADAS EN EL CÓDIGO (así solo necesitas subir el Excel)
# ──────────────────────────────────────────────────────────────────────────────
# Área de influencia (CONAF 2019, digitalizada de forma aproximada de la Fig. 40 del Plan de Manejo)
AREA_INFLUENCIA = json.loads('''{"type":"FeatureCollection","features":[{"type":"Feature","properties":{"nombre":"\u00c1rea de influencia PM CONAF 2019 (digitalizaci\u00f3n aproximada, Fig. 40)"},"geometry":{"type":"Polygon","coordinates":[[[-72.97935,-39.48796],[-72.98353,-39.49134],[-72.9834,-39.49575],[-72.97615,-39.50004],[-72.97751,-39.50227],[-72.98175,-39.50345],[-72.98593,-39.50683],[-72.9858,-39.51124],[-72.97709,-39.5166],[-72.97849,-39.51773],[-72.97836,-39.52214],[-72.97399,-39.52538],[-72.96822,-39.52748],[-72.96258,-39.52518],[-72.9584,-39.52179],[-72.9585,-39.51849],[-72.9529,-39.51508],[-72.953,-39.51177],[-72.95021,-39.50951],[-72.95034,-39.5051],[-72.95325,-39.50295],[-72.95332,-39.50074],[-72.94072,-39.4928],[-72.94089,-39.48729],[-72.94234,-39.48621],[-72.9381,-39.48503],[-72.93794,-39.49054],[-72.93211,-39.49485],[-72.92346,-39.49801],[-72.91783,-39.4957],[-72.90951,-39.48783],[-72.90965,-39.48342],[-72.91256,-39.48127],[-72.91269,-39.47685],[-72.90716,-39.47124],[-72.90729,-39.46683],[-72.9059,-39.4657],[-72.90302,-39.46675],[-72.89739,-39.46444],[-72.88911,-39.45547],[-72.88627,-39.45542],[-72.87905,-39.4586],[-72.87342,-39.45629],[-72.86925,-39.4529],[-72.86942,-39.44739],[-72.85833,-39.43726],[-72.85708,-39.43172],[-72.85155,-39.4261],[-72.8503,-39.42056],[-72.84332,-39.41602],[-72.83909,-39.41484],[-72.83475,-39.41696],[-72.8319,-39.41691],[-72.82607,-39.42121],[-72.81316,-39.42428],[-72.80737,-39.42748],[-72.80584,-39.43076],[-72.8056,-39.43848],[-72.79834,-39.44275],[-72.79685,-39.44493],[-72.79671,-39.44934],[-72.79233,-39.45257],[-72.78369,-39.45571],[-72.78785,-39.4591],[-72.78768,-39.46461],[-72.78469,-39.46897],[-72.78302,-39.47666],[-72.77865,-39.47989],[-72.77145,-39.48195],[-72.76996,-39.48413],[-72.77124,-39.48857],[-72.77402,-39.49083],[-72.77207,-39.50734],[-72.76911,-39.51059],[-72.76755,-39.51497],[-72.76028,-39.51925],[-72.76021,-39.52145],[-72.75287,-39.52793],[-72.75277,-39.53124],[-72.7569,-39.53573],[-72.75676,-39.54014],[-72.75376,-39.54449],[-72.75515,-39.54562],[-72.75498,-39.55113],[-72.75205,-39.55328],[-72.75052,-39.55656],[-72.75031,-39.56318],[-72.74592,-39.5664],[-72.7373,-39.56844],[-72.73006,-39.57161],[-72.72435,-39.5715],[-72.71163,-39.56794],[-72.69024,-39.56751],[-72.68885,-39.56638],[-72.68172,-39.56624],[-72.67737,-39.56836],[-72.67726,-39.57167],[-72.68713,-39.57517],[-72.69269,-39.5797],[-72.69675,-39.5864],[-72.70512,-39.59208],[-72.71928,-39.59567],[-72.73472,-39.60369],[-72.73904,-39.60267],[-72.74467,-39.60499],[-72.74884,-39.60838],[-72.7487,-39.61279],[-72.74578,-39.61494],[-72.74556,-39.62155],[-72.75106,-39.62828],[-72.75955,-39.63065],[-72.76101,-39.62957],[-72.7624,-39.6307],[-72.76387,-39.62963],[-72.76668,-39.63078],[-72.77104,-39.62866],[-72.77532,-39.62874],[-72.78256,-39.62557],[-72.78124,-39.62224],[-72.77846,-39.61998],[-72.7786,-39.61557],[-72.78298,-39.61234],[-72.78876,-39.61025],[-72.7944,-39.61256],[-72.79725,-39.61262],[-72.8016,-39.61049],[-72.80588,-39.61057],[-72.81152,-39.61289],[-72.82033,-39.60533],[-72.8261,-39.60323],[-72.82749,-39.60436],[-72.83038,-39.60331],[-72.83212,-39.59342],[-72.838,-39.58801],[-72.8366,-39.58688],[-72.83681,-39.58027],[-72.84119,-39.57704],[-72.84696,-39.57494],[-72.85124,-39.57502],[-72.86252,-39.57964],[-72.8793,-39.59098],[-72.88076,-39.5899],[-72.88789,-39.59003],[-72.89914,-39.59575],[-72.90783,-39.59149],[-72.91911,-39.59611],[-72.92346,-39.59398],[-72.93059,-39.59411],[-72.94943,-39.58451],[-72.9565,-39.58685],[-72.96068,-39.59023],[-72.9634,-39.59469],[-72.96765,-39.59587],[-72.97599,-39.60373],[-72.97573,-39.61255],[-72.97281,-39.61471],[-72.97385,-39.62797],[-72.96509,-39.63443],[-72.95351,-39.63974],[-72.95066,-39.63969],[-72.94916,-39.64187],[-72.94877,-39.6551],[-72.95017,-39.65623],[-72.95163,-39.65515],[-72.95302,-39.65628],[-72.95877,-39.65528],[-72.96442,-39.65758],[-72.9686,-39.66097],[-72.96841,-39.66758],[-72.96403,-39.67082],[-72.95393,-39.67395],[-72.94965,-39.67387],[-72.94961,-39.67497],[-72.95101,-39.6761],[-72.95088,-39.68051],[-72.95786,-39.68615],[-72.96653,-39.68299],[-72.97081,-39.68307],[-72.97646,-39.68537],[-72.98065,-39.68876],[-72.98052,-39.69317],[-72.97757,-39.69642],[-72.97897,-39.69755],[-72.97884,-39.70196],[-72.98303,-39.70534],[-72.9829,-39.70975],[-72.97998,-39.71191],[-72.98417,-39.71529],[-72.98391,-39.72411],[-72.98524,-39.72745],[-72.99103,-39.72534],[-72.99579,-39.70887],[-73.00018,-39.70564],[-73.00596,-39.70353],[-73.01167,-39.70363],[-73.01733,-39.70594],[-73.02152,-39.70932],[-73.02139,-39.71373],[-73.01993,-39.7148],[-73.01828,-39.7225],[-73.00951,-39.72897],[-73.00519,-39.73],[-73.00653,-39.73333],[-73.02635,-39.74029],[-73.03481,-39.74484],[-73.03627,-39.74376],[-73.04056,-39.74384],[-73.04479,-39.74611],[-73.05054,-39.74511],[-73.05619,-39.74741],[-73.06039,-39.75079],[-73.06027,-39.7552],[-73.05881,-39.75628],[-73.06167,-39.75633],[-73.06589,-39.7586],[-73.07025,-39.75647],[-73.07308,-39.75762],[-73.07737,-39.75769],[-73.07883,-39.75661],[-73.08306,-39.75889],[-73.08452,-39.75781],[-73.09018,-39.76011],[-73.09438,-39.76349],[-73.09423,-39.769],[-73.08111,-39.77761],[-73.07532,-39.77972],[-73.07103,-39.77965],[-73.0668,-39.77737],[-73.0538,-39.78157],[-73.04674,-39.77814],[-73.04385,-39.77919],[-73.04102,-39.77804],[-73.0367,-39.77907],[-73.02796,-39.78444],[-73.02646,-39.78662],[-73.0263,-39.79213],[-73.04459,-39.80347],[-73.04879,-39.80685],[-73.0487,-39.81016],[-73.0515,-39.81241],[-73.05433,-39.81356],[-73.06006,-39.81366],[-73.06444,-39.81042],[-73.07023,-39.80831],[-73.08296,-39.81404],[-73.08442,-39.81296],[-73.08582,-39.81409],[-73.09155,-39.81418],[-73.09301,-39.8131],[-73.09441,-39.81423],[-73.0973,-39.81317],[-73.10297,-39.81547],[-73.10312,-39.80996],[-73.1075,-39.80672],[-73.11329,-39.80461],[-73.11469,-39.80574],[-73.129,-39.80597],[-73.13189,-39.80491],[-73.13329,-39.80604],[-73.14045,-39.80615],[-73.14612,-39.80845],[-73.15032,-39.81183],[-73.15017,-39.81734],[-73.14579,-39.82058],[-73.14,-39.82269],[-73.13574,-39.82152],[-73.14833,-39.83275],[-73.15259,-39.83392],[-73.16101,-39.84068],[-73.16519,-39.84516],[-73.16808,-39.8441],[-73.16948,-39.84522],[-73.17664,-39.84534],[-73.19406,-39.83678],[-73.20265,-39.83692],[-73.21422,-39.83268],[-73.22275,-39.83502],[-73.22697,-39.83839],[-73.22533,-39.84609],[-73.22814,-39.84834],[-73.22803,-39.85275],[-73.22365,-39.85599],[-73.21786,-39.85811],[-73.21646,-39.85699],[-73.21064,-39.86021],[-73.20331,-39.86671],[-73.216,-39.87463],[-73.21583,-39.88125],[-73.22861,-39.88585],[-73.23566,-39.89037],[-73.23988,-39.89375],[-73.2412,-39.89818],[-73.24682,-39.90268],[-73.25972,-39.90287],[-73.26118,-39.90179],[-73.26548,-39.90186],[-73.26972,-39.90412],[-73.27402,-39.90419],[-73.27543,-39.90531],[-73.27835,-39.90315],[-73.2773,-39.88769],[-73.28169,-39.88445],[-73.28174,-39.88224],[-73.28612,-39.879],[-73.29191,-39.87688],[-73.3005,-39.877],[-73.30486,-39.87486],[-73.30916,-39.87492],[-73.31056,-39.87605],[-73.31494,-39.8728],[-73.32073,-39.87068],[-73.3264,-39.87297],[-73.33062,-39.87634],[-73.33051,-39.88075],[-73.33189,-39.88297],[-73.33476,-39.88301],[-73.34184,-39.88643],[-73.34758,-39.88651],[-73.35915,-39.88226],[-73.36209,-39.87899],[-73.36214,-39.87679],[-73.36801,-39.87135],[-73.36246,-39.86355],[-73.35824,-39.86018],[-73.35845,-39.85136],[-73.35991,-39.85028],[-73.35869,-39.84144],[-73.36449,-39.83821],[-73.36606,-39.83272],[-73.37043,-39.82947],[-73.37764,-39.82736],[-73.38056,-39.8252],[-73.38063,-39.82189],[-73.37639,-39.81962],[-73.37493,-39.82071],[-73.37066,-39.81954],[-73.36921,-39.82063],[-73.36637,-39.81948],[-73.36491,-39.82056],[-73.36064,-39.8194],[-73.35775,-39.82046],[-73.35208,-39.81818],[-73.34786,-39.81481],[-73.348,-39.80929],[-73.35237,-39.80605],[-73.35815,-39.80392],[-73.36674,-39.80404],[-73.36814,-39.80517],[-73.3696,-39.80408],[-73.37532,-39.80416],[-73.37957,-39.80643],[-73.38102,-39.80535],[-73.37821,-39.8031],[-73.37826,-39.8009],[-73.36697,-39.79412],[-73.36577,-39.78417],[-73.36874,-39.7798],[-73.36897,-39.76988],[-73.36754,-39.76986],[-73.35212,-39.7564],[-73.3522,-39.75309],[-73.34661,-39.7475],[-73.34677,-39.74088],[-73.34968,-39.73872],[-73.34976,-39.73541],[-73.35413,-39.73216],[-73.35294,-39.72221],[-73.34589,-39.7177],[-73.34603,-39.71219],[-73.34322,-39.70994],[-73.34483,-39.70224],[-73.34205,-39.69889],[-73.33208,-39.69765],[-73.32342,-39.70083],[-73.31776,-39.69854],[-73.31355,-39.69517],[-73.31363,-39.69186],[-73.30945,-39.68739],[-73.31112,-39.67749],[-73.30694,-39.67301],[-73.307,-39.67081],[-73.30422,-39.66746],[-73.29859,-39.66407],[-73.29875,-39.65745],[-73.29312,-39.65406],[-73.28903,-39.64627],[-73.28205,-39.63955],[-73.28068,-39.63733],[-73.28079,-39.63291],[-73.26953,-39.62613],[-73.26969,-39.61951],[-73.26692,-39.61616],[-73.26703,-39.61175],[-73.26423,-39.6095],[-73.26278,-39.61058],[-73.26138,-39.60946],[-73.25567,-39.60937],[-73.25001,-39.60708],[-73.24582,-39.60371],[-73.24033,-39.5948],[-73.24041,-39.59149],[-73.23625,-39.58702],[-73.23479,-39.5881],[-73.23197,-39.58695],[-73.23051,-39.58803],[-73.22768,-39.58689],[-73.22623,-39.58797],[-73.2234,-39.58682],[-73.22195,-39.5879],[-73.2163,-39.58561],[-73.2121,-39.58224],[-73.21219,-39.57893],[-73.2052,-39.5733],[-73.20531,-39.56889],[-73.20677,-39.56781],[-73.2054,-39.56558],[-73.20551,-39.56117],[-73.19707,-39.55663],[-73.19008,-39.55101],[-73.18592,-39.54653],[-73.18603,-39.54212],[-73.17893,-39.5409],[-73.17192,-39.53638],[-73.17067,-39.52974],[-73.1679,-39.52639],[-73.16811,-39.51867],[-73.16671,-39.51754],[-73.16685,-39.51203],[-73.16546,-39.5109],[-73.16261,-39.51086],[-73.16255,-39.51307],[-73.15237,-39.52063],[-73.14661,-39.52274],[-73.14091,-39.52265],[-73.13224,-39.52692],[-73.12659,-39.52463],[-73.12232,-39.52456],[-73.12086,-39.52564],[-73.11379,-39.52332],[-73.10806,-39.52433],[-73.0996,-39.52088],[-73.09545,-39.5164],[-73.08147,-39.50624],[-73.07452,-39.49951],[-73.07186,-39.49284],[-73.0677,-39.48836],[-73.06786,-39.48285],[-73.06931,-39.48177],[-73.06792,-39.48064],[-73.06807,-39.47513],[-73.08413,-39.46105],[-73.08134,-39.4588],[-73.08147,-39.45439],[-73.08292,-39.45331],[-73.07165,-39.44871],[-73.05769,-39.43855],[-73.05487,-39.4374],[-73.05054,-39.43953],[-73.04627,-39.43946],[-73.04042,-39.44488],[-73.04178,-39.44711],[-73.0416,-39.45373],[-73.03581,-39.45694],[-73.03575,-39.45914],[-73.03138,-39.46238],[-73.02562,-39.46449],[-73.02423,-39.46336],[-73.02135,-39.46441],[-73.0241,-39.46777],[-73.02398,-39.47218],[-73.01816,-39.47649],[-73.00664,-39.48071],[-73.001,-39.47841],[-72.99682,-39.47503],[-72.99698,-39.46951],[-72.99844,-39.46843],[-72.98717,-39.46383],[-72.98132,-39.46924],[-72.98125,-39.47145],[-72.98964,-39.47711],[-72.98951,-39.48152],[-72.97935,-39.48796]]]}}]}''')

# Polígono del santuario. Pega aquí tu GeoJSON (en lon/lat, EPSG:4326) entre las comillas triples.
# Cómo obtenerlo: entra a https://geojson.io, cambia el mapa a satélite, dibuja el polígono con la
# herramienta de polígono y copia el JSON del panel derecho. Déjalo vacío para no mostrarlo.
SANTUARIO_GEOJSON_TEXTO = ''''''

# Solo Valdivia y Mariquina (se puede apagar desde la barra lateral)
COMUNAS_PERMITIDAS = ["Valdivia", "Mariquina"]

# ── Tipo de proyecto → (icono Font Awesome 6, color del pin, etiqueta) ─────────
# Para cambiar un icono basta con editar este diccionario
# (catálogo: https://fontawesome.com/search?o=r&m=free)
TIPOS = {
    "Inmobiliario de borde":    {"icono": "fa-house",              "color": "#E67E22", "etiqueta": "Inmobiliario"},
    "Acuicultura":              {"icono": "fa-fish",               "color": "#1ABC9C", "etiqueta": "Acuicultura"},
    "Agroindustria":            {"icono": "fa-tractor",            "color": "#7D6608", "etiqueta": "Agroindustria"},
    "Combustibles":             {"icono": "fa-fire-flame-curved",  "color": "#C0392B", "etiqueta": "Combustibles"},
    "Drenaje":                  {"icono": "fa-grip-lines",         "color": "#5D6D7E", "etiqueta": "Drenaje"},
    "Forestal industrial":      {"icono": "fa-tree",               "color": "#1E8449", "etiqueta": "Forestal"},
    "Oro":                      {"icono": "fa-coins",              "color": "#B7950B", "etiqueta": "Oro"},
    "PTAS":                     {"icono": "fa-droplet",            "color": "#2E86C1", "etiqueta": "PTAS (plantas de tratamiento)"},
    "Planta Valdivia (Arauco)": {"icono": "fa-industry",           "color": "#34495E", "etiqueta": "Planta (Arauco Valdivia)"},
    "RILes":                    {"icono": "RILES",                 "color": "#117A65", "etiqueta": "RILes"},
    "Residuos":                 {"icono": "fa-dumpster",           "color": "#6E2C00", "etiqueta": "Residuos"},
    "Solar":                    {"icono": "fa-sun",                "color": "#D4AC0D", "etiqueta": "Solar"},
    "Tendido eléctrico":        {"icono": "fa-bolt",               "color": "#8E44AD", "etiqueta": "Tendido eléctrico"},
    "Áridos":                   {"icono": "fa-hammer",             "color": "#A04000", "etiqueta": "Áridos"},
    "Sin vía directa":          {"icono": "",                      "color": "#FF69B4", "etiqueta": "Sin vía directa"},
}

ESTADOS = ["Aprobado", "No Admitido a Tramitación", "Desistido", "No calificado", "Rechazado"]
COLOR_ESTADO = {
    "Aprobado": "#27AE60",
    "No Admitido a Tramitación": "#7F8C8D",
    "Desistido": "#E67E22",
    "No calificado": "#2980B9",
    "Rechazado": "#C0392B",
}

SIN_OBJETO = "Sin objeto afectado"
# Si quieres unificar nombres de objetos de conservación, agrégalos aquí. Ej:
# {"Fondos someros": "Fondos someros y totorales"}
NORMALIZAR_OBJETOS = {}

FA_CSS = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css"

# ──────────────────────────────────────────────────────────────────────────────
# 2. CARGA DE DATOS
# ──────────────────────────────────────────────────────────────────────────────
def ubicar(nombre: str) -> Path | None:
    """Busca el archivo en data/, junto a app.py, en la carpeta superior y, si no, en cualquier
    subcarpeta (acepta nombres con sufijo como 'archivo (1).xlsx')."""
    for carpeta in (BASE / "data", BASE, BASE.parent / "data", BASE.parent):
        ruta = carpeta / nombre
        if ruta.exists():
            return ruta
    raiz, ext = Path(nombre).stem, Path(nombre).suffix
    excluir = {".git", ".devcontainer", "node_modules", "__pycache__", ".venv"}
    for carpeta in (BASE, BASE.parent):
        for ruta in sorted(carpeta.rglob(f"{raiz}*{ext}")):
            if not excluir & set(ruta.parts):
                return ruta
    return None


def separar_objetos(texto) -> list[str]:
    """'Columna de agua, fondos someros (bajo)' -> ['Columna de agua', 'Fondos someros']"""
    if not isinstance(texto, str) or texto.strip() in ("", "—", "-"):
        return []
    texto = re.sub(r"\(.*?\)", "", texto)  # quita '(bajo)'
    salida = []
    for parte in texto.split(","):
        parte = parte.strip()
        if not parte:
            continue
        parte = parte[0].upper() + parte[1:]
        parte = NORMALIZAR_OBJETOS.get(parte, parte)
        if parte not in salida:
            salida.append(parte)
    return salida


@st.cache_data(show_spinner="Cargando proyectos…")
def cargar_proyectos(origen) -> pd.DataFrame:
    """origen: ruta (str) del Excel o los bytes de un Excel subido a mano."""
    if isinstance(origen, (bytes, bytearray)):
        origen = io.BytesIO(origen)
    df = pd.read_excel(origen, sheet_name=HOJA_PROYECTOS)
    df = df.rename(columns={
        "Nombre del proyecto": "nombre",
        "Tipo (clasificación del grupo)": "tipo",
        "Estado": "estado",
        "Titular": "titular",
        "Comuna": "comuna",
        "Latitud": "lat",
        "Longitud": "lon",
        "Vía hacia el santuario": "via",
        "Objetos de conservación afectados (PM, Tabla 17)": "objetos_txt",
        "Link expediente": "link",
    })
    df = df.dropna(subset=["lat", "lon"]).copy()
    for c in ("tipo", "estado", "comuna", "nombre", "titular", "via"):
        df[c] = df[c].astype(str).str.strip()
    df["objetos"] = df["objetos_txt"].apply(separar_objetos)
    return df


@st.cache_data
def cargar_geojson(nombre: str):
    ruta = ubicar(nombre)
    if ruta is None:
        return None
    with open(ruta, encoding="utf-8") as f:
        return json.load(f)


def _reproyectar(coords, transformador):
    if isinstance(coords[0], (int, float)):
        x, y = transformador.transform(coords[0], coords[1])
        return [x, y]
    return [_reproyectar(c, transformador) for c in coords]


def geojson_santuario():
    """Devuelve el GeoJSON del santuario (en lon/lat) o None si no existe el archivo.
    Si las coordenadas vienen en metros (UTM), las convierte desde EPSG:32719 (UTM 19S)."""
    gj = None
    for nombre in ARCHIVOS_SANTUARIO:
        gj = cargar_geojson(nombre)
        if gj is not None:
            break
    if gj is None and SANTUARIO_GEOJSON_TEXTO.strip():
        gj = json.loads(SANTUARIO_GEOJSON_TEXTO)
    if gj is None:
        return None
    muestra = json.dumps(gj)[:4000]
    crs = (gj.get("crs", {}).get("properties", {}).get("name", "") if isinstance(gj, dict) else "")
    primera = re.findall(r"-?\d+\.?\d*", re.sub(r'"crs".*?\}\}', "", muestra, flags=re.S))
    en_metros = any(abs(float(n)) > 360 for n in primera[:40] if "." in n or len(n) > 3)
    if en_metros:
        from pyproj import Transformer
        epsg = "EPSG:32718" if "32718" in crs else "EPSG:32719"
        t = Transformer.from_crs(epsg, "EPSG:4326", always_xy=True)
        for f in gj.get("features", []):
            f["geometry"]["coordinates"] = _reproyectar(f["geometry"]["coordinates"], t)
    return gj


# ──────────────────────────────────────────────────────────────────────────────
# 3. ICONOS (pines con Font Awesome)
# ──────────────────────────────────────────────────────────────────────────────
def html_icono(icono: str) -> str:
    """HTML del símbolo que va dentro del pin."""
    if icono == "RILES":  # basurero con gota de agua
        return (
            '<span class="fa-stack" style="font-size:9px;width:2.2em;height:2.2em;line-height:2.2em;">'
            '<i class="fa-solid fa-trash-can fa-stack-2x"></i>'
            '<i class="fa-solid fa-droplet fa-stack-1x" style="color:#5DADE2;font-size:.55em;'
            'top:.18em;"></i></span>'
        )
    if not icono:
        return ""
    return f'<i class="fa-solid {icono}"></i>'


def pin(tipo: str) -> folium.DivIcon:
    cfg = TIPOS[tipo]
    cuerpo = (
        f'<div style="position:relative;width:34px;height:42px;">'
        f'<div style="position:absolute;left:3px;top:0;width:28px;height:28px;background:{cfg["color"]};'
        f'border:2px solid #fff;border-radius:50% 50% 50% 0;transform:rotate(-45deg);'
        f'box-shadow:0 2px 5px rgba(0,0,0,.45);"></div>'
        f'<div style="position:absolute;left:3px;top:0;width:28px;height:28px;display:flex;'
        f'align-items:center;justify-content:center;color:#fff;font-size:13px;">'
        f'{html_icono(cfg["icono"])}</div></div>'
    )
    return folium.DivIcon(html=cuerpo, icon_size=(34, 42), icon_anchor=(17, 40), popup_anchor=(0, -36))


def popup_html(r) -> str:
    esc = html.escape
    color = COLOR_ESTADO.get(r["estado"], "#555")
    objetos = ", ".join(r["objetos"]) if r["objetos"] else "Sin objeto afectado"
    link = ""
    if isinstance(r.get("link"), str) and r["link"].startswith("http"):
        link = f'<a href="{esc(r["link"])}" target="_blank">Ver expediente SEIA ↗</a>'
    return f"""
    <div style="font-family:Arial,sans-serif;font-size:12.5px;width:310px;line-height:1.4;">
      <div style="font-weight:700;font-size:13.5px;margin-bottom:6px;">{esc(r["nombre"])}</div>
      <div><b>Titular:</b> {esc(r["titular"])}</div>
      <div><b>Comuna:</b> {esc(r["comuna"])}</div>
      <div><b>Vía hacia el santuario:</b> {esc(r["via"])}</div>
      <hr style="margin:6px 0;border:0;border-top:1px solid #ddd;">
      <div><b>Tipo:</b> {esc(TIPOS.get(r["tipo"], {}).get("etiqueta", r["tipo"]))}</div>
      <div><b>Estado:</b> <span style="color:{color};font-weight:700;">{esc(r["estado"])}</span></div>
      <div><b>Objetos afectados:</b> {esc(objetos)}</div>
      <div style="margin-top:4px;">{link}</div>
    </div>"""


# ──────────────────────────────────────────────────────────────────────────────
# 4. ESTADO DE LOS FILTROS (leyendas con checkbox)
# ──────────────────────────────────────────────────────────────────────────────
def fijar_todos(prefijo: str, claves: list[str], valor: bool):
    for c in claves:
        st.session_state[f"{prefijo}_{c}"] = valor


def leyenda_checks(prefijo: str, claves: list[str], etiquetas=None, iconos=None, notas=None, conteos=None,
                   defecto: bool = True):
    """Dibuja una leyenda con checkboxes y devuelve las claves marcadas."""
    c1, c2 = st.columns(2)
    c1.button("✔ Marcar todo", key=f"{prefijo}_on", width="stretch",
              on_click=fijar_todos, args=(prefijo, claves, True))
    c2.button("✖ Desmarcar", key=f"{prefijo}_off", width="stretch",
              on_click=fijar_todos, args=(prefijo, claves, False))
    marcadas = []
    for c in claves:
        k = f"{prefijo}_{c}"
        if k not in st.session_state:
            st.session_state[k] = defecto
        col_i, col_c = st.columns([1, 8], vertical_alignment="center")
        if iconos:
            col_i.markdown(iconos[c], unsafe_allow_html=True)
        n = f" ({conteos.get(c, 0)})" if conteos is not None else ""
        etiqueta = (etiquetas or {}).get(c, c) + n
        if st.checkbox(etiqueta, key=k):
            marcadas.append(c)
        if notas and notas.get(c):
            st.caption(notas[c])
    return marcadas


def mini_pin_html(tipo: str) -> str:
    cfg = TIPOS[tipo]
    return (
        f'<span style="display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;'
        f'background:{cfg["color"]};color:#fff;border-radius:50%;font-size:10px;">{html_icono(cfg["icono"])}</span>'
    )


def punto_estado_html(estado: str) -> str:
    return (f'<span style="display:inline-block;width:14px;height:14px;border-radius:50%;'
            f'background:{COLOR_ESTADO[estado]};"></span>')


# ──────────────────────────────────────────────────────────────────────────────
# 5. APLICACIÓN
# ──────────────────────────────────────────────────────────────────────────────
st.markdown(f"<style>@import url('{FA_CSS}');</style>", unsafe_allow_html=True)
st.markdown(
    "<style>.block-container{padding-top:1.2rem;} "
    "[data-testid='stSidebar'] .stCheckbox{margin-bottom:-0.6rem;}</style>",
    unsafe_allow_html=True,
)

ruta_excel = ubicar(ARCHIVO_EXCEL)
if ruta_excel is not None:
    df_total = cargar_proyectos(str(ruta_excel))
else:
    st.warning(f"No encuentro **{ARCHIVO_EXCEL}** en el repositorio. Súbelo aquí para usar la app "
               "(luego agrégalo a la carpeta de `app.py` para no repetir este paso).")
    subido = st.file_uploader("Excel de proyectos (.xlsx)", type=["xlsx"])
    if subido is None:
        st.stop()
    df_total = cargar_proyectos(subido.getvalue())

st.title("🦢 Santuario de la Naturaleza Río Cruces")
st.caption("Proyectos del SEIA en el área de influencia · Valdivia y San José de la Mariquina, Región de Los Ríos")

# ── Opciones generales ────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Filtros")
    st.caption("Las tres leyendas se **suman**: un proyecto aparece si cumple alguna casilla marcada "
               "en Tipos, en Estado o en Objetos.")
    solo_vm = st.toggle("Solo Valdivia y Mariquina", value=True)
    base_elegida = st.radio("Mapa base", ["OpenStreetMap", "OpenStreetMap (alternativo)", "Satélite (Esri)"],
                            help="Si OpenStreetMap se ve vacío, prueba el alternativo (mismo mapa, otro servidor).")
    agrupar = st.toggle("Agrupar pines cercanos", value=True,
                        help="Hay proyectos con coordenada genérica repetida; al acercar el mapa se separan.")

df = df_total.copy()
if solo_vm:
    patron = "|".join(COMUNAS_PERMITIDAS)
    df = df[df["comuna"].str.contains(patron, case=False, na=False)]

# Objetos de conservación presentes en los datos
objetos_unicos = sorted({o for lista in df["objetos"] for o in lista})
claves_objetos = objetos_unicos + [SIN_OBJETO]

with st.sidebar:
    tab_tipo, tab_estado, tab_obj = st.tabs(["🧭 Tipos", "📋 Estado", "🌿 Objetos"])

    # --- Leyenda 1: tipos de proyecto -----------------------------------------
    with tab_tipo:
        orden = [t for t in TIPOS if t in set(df["tipo"])]
        tipos_marcados = leyenda_checks(
            "tipo", orden,
            etiquetas={t: TIPOS[t]["etiqueta"] for t in orden},
            iconos={t: mini_pin_html(t) for t in orden},
            conteos=df["tipo"].value_counts().to_dict(),
        )

    # --- Leyenda 2: estado de tramitación -------------------------------------
    with tab_estado:
        estados_presentes = [e for e in ESTADOS if e in set(df["estado"])]
        estados_marcados = leyenda_checks(
            "estado", estados_presentes,
            iconos={e: punto_estado_html(e) for e in estados_presentes},
            conteos=df["estado"].value_counts().to_dict(),
            defecto=False,
        )

    # --- Leyenda 3: objetos de conservación afectados -------------------------
    with tab_obj:
        st.caption("Objetos de conservación del Plan de Manejo (Tabla 17) que afecta cada proyecto. "
                   "Debajo de cada uno se indican los tipos de proyecto que lo afectan.")
        notas, conteo_obj = {}, {}
        for o in objetos_unicos:
            sub = df[df["objetos"].apply(lambda lista, o=o: o in lista)]
            conteo_obj[o] = len(sub)
            tipos_o = sorted(TIPOS[t]["etiqueta"] for t in sub["tipo"].unique() if t in TIPOS)
            notas[o] = "Afectado por: " + ", ".join(tipos_o)
        sub_sin = df[df["objetos"].apply(len) == 0]
        conteo_obj[SIN_OBJETO] = len(sub_sin)
        notas[SIN_OBJETO] = "Proyectos dentro del área CONAF sin vía directa al agua."
        objetos_marcados = leyenda_checks(
            "obj", claves_objetos, notas=notas, conteos=conteo_obj, defecto=False,
        )

# ── Aplicar filtros: cada leyenda es independiente y se suman (unión) ─────────
def pasa_objetos(lista) -> bool:
    if not lista:
        return SIN_OBJETO in objetos_marcados
    return any(o in objetos_marcados for o in lista)


filtrado = df[
    df["tipo"].isin(tipos_marcados)
    | df["estado"].isin(estados_marcados)
    | df["objetos"].apply(pasa_objetos)
]

m1, m2, m3 = st.columns(3)
m1.metric("Proyectos mostrados", len(filtrado))
m2.metric("Aprobados", int((filtrado["estado"] == "Aprobado").sum()))
m3.metric("Tipos activos", filtrado["tipo"].nunique())

# ── Mapa ──────────────────────────────────────────────────────────────────────
mapa = folium.Map(location=CENTRO_SANTUARIO, zoom_start=ZOOM_INICIAL, tiles=None, control_scale=True)
mapa.get_root().header.add_child(folium.CssLink(FA_CSS))

# Capas base: OpenStreetMap y satélite
# (referrer_policy evita que OSM bloquee los mosaicos dentro del mapa incrustado de Streamlit)
folium.TileLayer(
    tiles="https://tile.openstreetmap.org/{z}/{x}/{y}.png",
    attr="&copy; Colaboradores de OpenStreetMap",
    name="OpenStreetMap",
    max_zoom=19,
    show=(base_elegida == "OpenStreetMap"),
    referrer_policy="strict-origin-when-cross-origin",
).add_to(mapa)
folium.TileLayer(
    tiles="https://{s}.tile.openstreetmap.fr/osmfr/{z}/{x}/{y}.png",
    attr="&copy; OpenStreetMap France | &copy; Colaboradores de OpenStreetMap",
    name="OpenStreetMap (alternativo)",
    max_zoom=20,
    show=(base_elegida == "OpenStreetMap (alternativo)"),
).add_to(mapa)
folium.TileLayer(
    tiles="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
    attr="Esri, Maxar, Earthstar Geographics",
    name="Satélite (Esri)",
    max_zoom=19,
    show=(base_elegida == "Satélite (Esri)"),
).add_to(mapa)

# Área de influencia (morado oscuro)
influencia = cargar_geojson(ARCHIVO_INFLUENCIA) or AREA_INFLUENCIA
if influencia:
    folium.GeoJson(
        influencia,
        name="Área de influencia (CONAF 2019, aprox.)",
        style_function=lambda _: {"color": COLOR_INFLUENCIA, "weight": 2.5,
                                  "fillColor": COLOR_INFLUENCIA, "fillOpacity": 0.18},
        tooltip="Área de influencia (CONAF 2019, aproximada)",
    ).add_to(mapa)
else:
    st.sidebar.warning(f"Falta data/{ARCHIVO_INFLUENCIA}")

# Santuario (celeste claro)
gj_santuario = geojson_santuario()
if gj_santuario is not None:
    folium.GeoJson(
        gj_santuario,
        name="Santuario Río Cruces",
        style_function=lambda _: {"color": "#3C9BD6", "weight": 2.5,
                                  "fillColor": COLOR_SANTUARIO, "fillOpacity": 0.55},
        tooltip="Santuario de la Naturaleza Río Cruces",
    ).add_to(mapa)
else:
    st.sidebar.warning("Falta el polígono del santuario. Pega tu GeoJSON en `SANTUARIO_GEOJSON_TEXTO` "
                       "(dentro de app.py) o sube `santuario_rio_cruces.geojson`.")

# Pines
contenedor = MarkerCluster(name="Proyectos", options={"disableClusteringAtZoom": 14,
                                                       "spiderfyOnMaxZoom": True}) if agrupar \
    else folium.FeatureGroup(name="Proyectos")
for _, r in filtrado.iterrows():
    if r["tipo"] not in TIPOS:
        continue
    folium.Marker(
        [r["lat"], r["lon"]],
        icon=pin(r["tipo"]),
        tooltip=r["nombre"][:90],
        popup=folium.Popup(popup_html(r), max_width=340),
    ).add_to(contenedor)
contenedor.add_to(mapa)

Fullscreen().add_to(mapa)
folium.LayerControl(collapsed=True, position="topright").add_to(mapa)

st_folium(mapa, height=680, use_container_width=True, returned_objects=[])

with st.expander("📄 Tabla de proyectos mostrados"):
    tabla = filtrado[["nombre", "titular", "comuna", "tipo", "estado", "via", "objetos_txt"]].rename(columns={
        "nombre": "Proyecto", "titular": "Titular", "comuna": "Comuna", "tipo": "Tipo",
        "estado": "Estado", "via": "Vía hacia el santuario", "objetos_txt": "Objetos afectados"})
    st.dataframe(tabla, width="stretch", hide_index=True)

st.caption("Fuentes: SEIA (seia.sea.gob.cl); Plan de Manejo MN y SN Río Cruces y Chorocamayo, CONAF 2019. "
           "Los puntos son la coordenada representativa del SEIA, no necesariamente la faena. "
           "El área de influencia está digitalizada de forma aproximada: solo para exploración.")
