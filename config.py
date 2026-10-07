"""Todas as configuracoes do pipeline de trilhas.

Variaveis de ambiente sobrescrevem os valores padrao quando existirem.
"""
import os
import re
from datetime import datetime, timezone
from pathlib import Path

# ---------- Pastas base ----------
DATA = Path("data")
UTILS_DIR = DATA / "utils"  # DEM (.tif/.vrt) e cache bruto do Overpass

# ---------- Variaveis de ambiente ----------
RUN_DATE = datetime.now(timezone.utc).strftime("%Y%m%d")
BBOX = os.getenv("BBOX", "-24.05,-47.10,-23.00,-46.20")
GAP_M = float(os.getenv("GAP_M", "30"))
NAME_REGEX = os.getenv("NAME_REGEX", "trilha|caminho|pico|cachoeira|pedra")
NAME_RE = re.compile(NAME_REGEX, re.IGNORECASE) if NAME_REGEX else None
HIGHWAYS = [
    h.strip()
    for h in os.getenv("HIGHWAYS", "path,footway,track").split(",")
    if h.strip()
]
REFRESH = os.getenv("REFRESH", "0") == "1"
ELEVATION = os.getenv("ELEVATION", "1") == "1"
DEM_FILE = os.getenv("DEM_FILE", str(UTILS_DIR / "copernicus.vrt"))  # Arquivo DEM offline

# ---------- Parametros de calculo ----------
SMOOTH_RADIUS = 2      # media movel de pontos
HYSTERESIS_M = 4.0     # ignora oscilacoes menores que isso
CIRCULAR_M = 100.0     # inicio ~ fim => circular
EASY_MAX = 6.0
MODERATE_MAX = 14.0
EFFORT_DESNIVEL_DIV = 100   # metros de desnivel equivalentes a 1 km de esforco
WALK_SPEED_KMH = 5          # velocidade horizontal para estimativa de tempo
CLIMB_SPEED_M_H = 600       # velocidade vertical (m/h) para estimativa de tempo

# ---------- Caminhos de saida ----------
RAW_GLOB = "overpass_raw_*.json"
XLSX_DIR = DATA / "xlsx"
GEOJSON_DIR = DATA / "geojson"
SQLITE_DIR = DATA / "sqlite"
GPX_ROOT = DATA / "gpx"

OUT_EXCEL = XLSX_DIR / f"trilhas_{RUN_DATE}.xlsx"
OUT_GEOJSON = GEOJSON_DIR / f"trilhas_{RUN_DATE}.geojson"
OUT_SQLITE = SQLITE_DIR / f"trilhas_{RUN_DATE}.sqlite"
GPX_DIR = GPX_ROOT / f"gpx_{RUN_DATE}"  # um arquivo .gpx por trilha

# ---------- Overpass API ----------
ENDPOINTS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
]
HEADERS = {"User-Agent": "trilhas-sp-spike/0.2"}
OVERPASS_QUERY_TIMEOUT_S = 180   # timeout declarado dentro da query
REQUEST_TIMEOUT_S = 240          # timeout do requests.post
MAX_ATTEMPTS = 3
RETRY_SLEEP_S = 5                # pausa entre endpoints
BACKOFF_S = 15                   # pausa entre tentativas (multiplicada pela tentativa)
QUERY_PAUSE_S = 5                # pausa entre queries de highways diferentes

# ---------- Tags OSM ----------
EXTRA_TAGS = [
    "surface", "sac_scale", "trail_visibility", "access", "foot", "bicycle",
    "horse", "dog", "wheelchair", "incline", "width", "operator", "website",
    "description", "ref", "wikidata", "wikipedia",
]
SAC_ORDER = [
    "hiking",
    "mountain_hiking",
    "demanding_mountain_hiking",
    "alpine_hiking",
    "demanding_alpine_hiking",
    "difficult_alpine_hiking",
]

# ---------- Colunas de saida ----------
SQLITE_BASE_COLUMNS = [
    "trilha_id TEXT PRIMARY KEY",
    "nome TEXT",
    "nome_bate_regex TEXT",
    "km_principal REAL",
    "km_total REAL",
    "n_ways INTEGER",
    "segmentos INTEGER",
    "inicio_lat REAL",
    "inicio_lon REAL",
    "fim_lat REAL",
    "fim_lon REAL",
    "circular TEXT",
    "dificuldade_calc TEXT",
    "fonte_dificuldade TEXT",
    "elevacao_status TEXT",
    "desnivel_acumulado REAL",
    "ganho_m REAL",
    "perda_m REAL",
    "tempo_est_calc TEXT",
    "ultima_edicao TEXT",
    "highway TEXT",
]

EXCEL_HEADERS = [
    "trilha_id", "nome", "nome_bate_regex", "km_principal", "km_total",
    "n_ways", "segmentos", "inicio_lat", "inicio_lon", "fim_lat", "fim_lon",
    "circular", "dificuldade_calc", "fonte_dificuldade", "elevacao_status",
    "desnivel_acumulado", "ganho_m", "perda_m", "tempo_est_calc",
    "ultima_edicao", "highway", *EXTRA_TAGS,
]