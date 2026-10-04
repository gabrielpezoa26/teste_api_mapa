"""Overpass -> junta ways por nome -> ordena -> elevacao -> dificuldade -> Excel + GeoJSON.

Objetivo: base BRUTA. Nada e descartado na coleta (todos os ways nomeados de path/footway/track);
regex de nome, tamanho etc. viram colunas-marcador para filtrar depois.

Juncao de ways do mesmo nome:
  1. por node compartilhado (conexao exata do OSM)
  2. por proximidade das pontas (GAP_M metros)
Depois os ways de cada trilha sao encadeados em linhas continuas.
"""
import atexit
import json
import math
import os
import re
import sys
import time
import unicodedata
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

BBOX = os.getenv("BBOX", "-24.05,-47.10,-23.00,-46.20")
GAP_M = float(os.getenv("GAP_M", "30"))
NAME_REGEX = os.getenv("NAME_REGEX", "trilha|caminho|pico|cachoeira|pedra")  # so MARCA (nome_bate_regex)
NAME_RE = re.compile(NAME_REGEX, re.IGNORECASE) if NAME_REGEX else None
HIGHWAYS = [h.strip() for h in os.getenv("HIGHWAYS", "path,footway,track").split(",") if h.strip()]
REFRESH = os.getenv("REFRESH", "0") == "1"
ELEVATION = os.getenv("ELEVATION", "1") == "1"
STEP_M = float(os.getenv("STEP_M", "50"))
MIN_KM_ELEV = float(os.getenv("MIN_KM_ELEV", "0"))

MAX_SAMPLES = 400      # maximo de pontos de elevacao por trilha
SMOOTH_RADIUS = 2      # media movel de 5 pontos
HYSTERESIS_M = 4.0     # ignora oscilacoes menores que isso
CIRCULAR_M = 100.0     # inicio ~ fim => circular
EASY_MAX, MODERATE_MAX = 6.0, 14.0   # limiares de esforco (km + ganho_m/100)

DATA = Path("data")
RAW_GLOB = "overpass_raw_*.json"  # um arquivo por extracao: overpass_raw_AAAAMMDD_HHMMSS.json
ELEV_CACHE = DATA / "elevation_cache.json"
OUT = DATA / "trilhas.xlsx"
OUT_GEOJSON = DATA / "trilhas.geojson"

ENDPOINTS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.private.coffee/api/interpreter",
]
HEADERS = {"User-Agent": "trilhas-sp-spike/0.1 (validacao de dados; contato: troque-aqui)"}
ELEV_URL = "https://api.open-meteo.com/v1/elevation"

EXTRA_TAGS = ["surface", "sac_scale", "trail_visibility", "access", "foot", "bicycle",
              "horse", "dog", "wheelchair", "incline", "width", "operator", "website",
              "description", "ref", "wikidata", "wikipedia"]
SAC_ORDER = ["hiking", "mountain_hiking", "demanding_mountain_hiking",
             "alpine_hiking", "demanding_alpine_hiking", "difficult_alpine_hiking"]


def fix_owner():
    """Devolve data/ ao dono da pasta do projeto (evita arquivos de root)."""
    try:
        st = os.stat(".")
        for f in [DATA, *DATA.glob("*")]:
            os.chown(f, st.st_uid, st.st_gid)
    except (OSError, AttributeError):
        pass


atexit.register(fix_owner)


# ---------- 1. Overpass ----------
def build_query(highway: str) -> str:
    return (
        "[out:json][timeout:180];\n"
        f'way[highway="{highway}"][name]({BBOX});\n'
        "out meta geom;"  # nodes + coordenadas + version/timestamp
    )


def post_overpass(query: str):
    """Tenta os endpoints com retry. Devolve (json, endpoint)."""
    last_err = None
    for attempt in range(1, 4):
        for url in ENDPOINTS:
            try:
                print(f"  [{attempt}/3] {url} ...")
                r = requests.post(url, data={"data": query}, headers=HEADERS, timeout=240)
                if r.status_code == 200:
                    return r.json(), url
                last_err = f"HTTP {r.status_code}"
            except (requests.RequestException, ValueError) as e:
                last_err = str(e)
            print(f"    falhou: {last_err}")
            time.sleep(5)
        time.sleep(15 * attempt)
    sys.exit(f"Overpass indisponivel: {last_err}. Tente de novo mais tarde.")


def fetch() -> tuple[dict, dict]:
    """Retorna (dados, meta). Usa o cache mais recente se BBOX e HIGHWAYS forem os mesmos."""
    DATA.mkdir(exist_ok=True)
    files = sorted(DATA.glob(RAW_GLOB))
    if files and not REFRESH:
        saved = json.loads(files[-1].read_text(encoding="utf-8"))
        meta = saved.get("meta", {})
        if meta.get("bbox") == BBOX and meta.get("highways") == HIGHWAYS:
            meta["cache_file"] = files[-1].name
            print(f"Usando cache {files[-1]} (REFRESH=1 para refazer)")
            return saved, meta
        print("Cache de outra BBOX/HIGHWAYS: consultando de novo.")

    elements, endpoints, osm_bases = {}, {}, []
    for i, hw in enumerate(HIGHWAYS):
        if i:
            time.sleep(5)  # uma consulta por vez, com pausa (boa pratica do Overpass)
        query = build_query(hw)
        print(f"Query highway={hw}:\n{query}")
        data, url = post_overpass(query)
        endpoints[hw] = url
        base = data.get("osm3s", {}).get("timestamp_osm_base", "")
        if base:
            osm_bases.append(base)
        for el in data.get("elements", []):
            elements[(el.get("type"), el.get("id"))] = el
        print(f"  {len(data.get('elements', []))} elementos")

    now = datetime.now(timezone.utc)
    meta = {
        "extracted_at": now.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "osm_base": min(osm_bases) if osm_bases else "",
        "bbox": BBOX,
        "highways": HIGHWAYS,
        "endpoints": endpoints,
    }
    saved = {"meta": meta, "elements": list(elements.values())}
    out = DATA / f"overpass_raw_{now.strftime('%Y%m%d_%H%M%S')}.json"
    out.write_text(json.dumps(saved), encoding="utf-8")
    print(f"Extracao salva em {out}")
    meta["cache_file"] = out.name
    return saved, meta


# ---------- 2. Geometria ----------
def haversine(a, b) -> float:
    (lat1, lon1), (lat2, lon2) = a, b
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi, dl = p2 - p1, math.radians(lon2 - lon1)
    h = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * 6371000 * math.asin(math.sqrt(h))


def line_length(coords) -> float:
    return sum(haversine(coords[i], coords[i + 1]) for i in range(len(coords) - 1))


def norm(name: str) -> str:
    s = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s.lower()).strip()


class UF:
    def __init__(self, n):
        self.p = list(range(n))

    def find(self, x):
        while self.p[x] != x:
            self.p[x] = self.p[self.p[x]]
            x = self.p[x]
        return x

    def union(self, a, b):
        self.p[self.find(a)] = self.find(b)


def chain_ways(ways: list[dict]):
    """Encadeia ways em linhas continuas (gulosamente, pelas pontas mais proximas).
    Retorna (segmentos, maior_lacuna_m). Segmentos ordenados do maior para o menor;
    mais de 1 segmento = ramificacao ou pedaco que nao encaixou dentro de GAP_M."""
    remaining = sorted((list(w["coords"]) for w in ways), key=line_length, reverse=True)
    segments, max_gap = [], 0.0
    while remaining:
        cur = remaining.pop(0)
        for end in ("tail", "head"):
            while True:
                point = cur[-1] if end == "tail" else cur[0]
                best = None  # (dist, indice, inverter?)
                for i, c in enumerate(remaining):
                    for first, inv in ((c[0], False), (c[-1], True)):
                        d = haversine(point, first)
                        if d <= GAP_M and (best is None or d < best[0]):
                            best = (d, i, inv)
                if best is None:
                    break
                d, i, inv = best
                piece = remaining.pop(i)
                if end == "tail":
                    piece = piece[::-1] if inv else piece
                    cur = cur + piece
                else:
                    piece = piece if inv else piece[::-1]
                    cur = piece + cur
                max_gap = max(max_gap, d)
        segments.append(cur)
    segments.sort(key=line_length, reverse=True)
    return segments, max_gap


# ---------- 3. Juncao por nome ----------
def parse_ways(data: dict) -> list[dict]:
    ways = []
    for el in data.get("elements", []):
        if el.get("type") != "way" or not el.get("geometry"):
            continue
        tags = el.get("tags", {})
        coords = [(p["lat"], p["lon"]) for p in el["geometry"]]
        ways.append({
            "id": el["id"],
            "name": tags.get("name", ""),
            "key": norm(tags.get("name", "")),
            "highway": tags.get("highway", ""),
            "tags": tags,
            "timestamp": el.get("timestamp", ""),
            "version": el.get("version", 0),
            "nodes": el.get("nodes", []),
            "coords": coords,
            "length_m": line_length(coords),
        })
    return ways


def distinct(ws, key):
    return sorted({w["tags"][key] for w in ws if w["tags"].get(key)})


def merge_group(group: list[dict]) -> list[dict]:
    n = len(group)
    uf_nodes = UF(n)
    owner = {}
    for i, w in enumerate(group):
        for nid in w["nodes"]:
            if nid in owner:
                uf_nodes.union(i, owner[nid])
            else:
                owner[nid] = i
    node_root = [uf_nodes.find(i) for i in range(n)]

    uf_gap = UF(n)
    for i in range(n):
        uf_gap.union(i, node_root[i])
    ends = []
    for i, w in enumerate(group):
        ends.append((w["coords"][0], i))
        ends.append((w["coords"][-1], i))
    for a in range(len(ends)):
        for b in range(a + 1, len(ends)):
            ia, ib = ends[a][1], ends[b][1]
            if uf_gap.find(ia) == uf_gap.find(ib):
                continue
            if haversine(ends[a][0], ends[b][0]) <= GAP_M:
                uf_gap.union(ia, ib)

    clusters = defaultdict(list)
    for i in range(n):
        clusters[uf_gap.find(i)].append(i)

    result = []
    for idxs in clusters.values():
        ws = [group[i] for i in idxs]
        pts = [p for w in ws for p in w["coords"]]
        stamps = [w["timestamp"] for w in ws if w["timestamp"]]
        all_tags = {}
        for w in ws:
            for k, v in w["tags"].items():
                all_tags.setdefault(k, set()).add(v)
        segments, max_gap = chain_ways(ws)
        main = segments[0]
        t = {
            "name": ws[0]["name"],
            "name_matches": "sim" if (NAME_RE is None or NAME_RE.search(ws[0]["name"])) else "nao",
            "n_ways": len(ws),
            "pieces_by_node": len({node_root[i] for i in idxs}),
            "segments": segments,
            "n_segments": len(segments),
            "max_gap_m": round(max_gap),
            "length_km": round(sum(w["length_m"] for w in ws) / 1000, 2),
            "lat": round(sum(p[0] for p in pts) / len(pts), 5),
            "lon": round(sum(p[1] for p in pts) / len(pts), 5),
            "start": main[0],
            "end": main[-1],
            "circular": haversine(main[0], main[-1]) < CIRCULAR_M,
            "highway": ", ".join(sorted({w["highway"] for w in ws if w["highway"]})),
            "last_edit": max(stamps) if stamps else "",
            "first_edit": min(stamps) if stamps else "",
            "max_version": max(w["version"] for w in ws),
            "tags_json": json.dumps({k: sorted(v) for k, v in sorted(all_tags.items())},
                                    ensure_ascii=False)[:32000],
            "way_ids": ", ".join(str(w["id"]) for w in ws),
            "elev": None,
            "elev_status": "desativada",
        }
        for k in EXTRA_TAGS:
            t[k] = ", ".join(distinct(ws, k))
        result.append(t)
    return result


# ---------- 4. Elevacao (Open-Meteo, DEM de 90 m) ----------
def ekey(lat, lon) -> str:
    return f"{lat:.5f},{lon:.5f}"


def resample(coords, step_m):
    """Pontos a cada step_m ao longo da linha: [(dist_m, lat, lon), ...]."""
    out = [(0.0, coords[0][0], coords[0][1])]
    need, base = step_m, 0.0
    for a, b in zip(coords, coords[1:]):
        seg = haversine(a, b)
        if seg == 0:
            continue
        pos = 0.0
        while seg - pos >= need:
            pos += need
            t = pos / seg
            out.append((base + pos, a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t))
            need = step_m
        need -= seg - pos
        base += seg
    if out[-1][1:] != tuple(coords[-1]):
        out.append((base, coords[-1][0], coords[-1][1]))
    return out


def fetch_elevations(keys: list[str], cache: dict) -> None:
    missing = [k for k in keys if k not in cache]
    print(f"Elevacao: {len(keys)} pontos, {len(missing)} a buscar "
          f"({math.ceil(len(missing) / 100)} requisicoes)")
    for i in range(0, len(missing), 100):
        batch = missing[i:i + 100]
        lats = ",".join(k.split(",")[0] for k in batch)
        lons = ",".join(k.split(",")[1] for k in batch)
        for attempt in range(1, 6):
            try:
                r = requests.get(ELEV_URL, params={"latitude": lats, "longitude": lons},
                                 headers=HEADERS, timeout=60)
                if r.status_code == 200:
                    for k, e in zip(batch, r.json()["elevation"]):
                        cache[k] = e
                    break
                print(f"  Open-Meteo HTTP {r.status_code} (tentativa {attempt})")
            except (requests.RequestException, ValueError, KeyError) as e:
                print(f"  Open-Meteo erro: {e} (tentativa {attempt})")
            time.sleep(10 * attempt)
        else:
            print("  desistindo do lote; trilhas sem elevacao completa ficam em branco")
            break
        ELEV_CACHE.write_text(json.dumps(cache), encoding="utf-8")
        time.sleep(1)  # respeita o limite da API gratuita
    ELEV_CACHE.write_text(json.dumps(cache), encoding="utf-8")


def smooth(vals, r=SMOOTH_RADIUS):
    return [sum(vals[max(0, i - r):i + r + 1]) / len(vals[max(0, i - r):i + r + 1])
            for i in range(len(vals))]


def gain_loss(vals, thr=HYSTERESIS_M):
    ref, gain, loss = vals[0], 0.0, 0.0
    for e in vals[1:]:
        d = e - ref
        if d >= thr:
            gain += d
            ref = e
        elif d <= -thr:
            loss += -d
            ref = e
    return gain, loss


def add_elevation(trails: list[dict]) -> None:
    cache = json.loads(ELEV_CACHE.read_text()) if ELEV_CACHE.exists() else {}
    plans, keys = {}, set()
    for idx, t in enumerate(trails):
        if t["length_km"] < MIN_KM_ELEV:
            t["elev_status"] = "pulada"
            continue
        total = sum(line_length(s) for s in t["segments"])
        step = max(STEP_M, total / MAX_SAMPLES)
        plan = [resample(s, step) for s in t["segments"]]
        plans[idx] = (plan, [line_length(s) for s in t["segments"]])
        keys.update(ekey(p[1], p[2]) for seg in plan for p in seg)
    fetch_elevations(sorted(keys), cache)

    for idx, (plan, seg_lens) in plans.items():
        gain = loss = 0.0
        all_el, profile, offset = [], [], 0.0
        ok = True
        for seg, seg_len in zip(plan, seg_lens):
            vals = [cache.get(ekey(p[1], p[2])) for p in seg]
            if any(v is None for v in vals):
                ok = False
                break
            sm = smooth(vals)
            g, l = gain_loss(sm)
            gain, loss = gain + g, loss + l
            all_el += sm
            profile += [(offset + p[0], e) for p, e in zip(seg, sm)]
            offset += seg_len
        if not ok:
            trails[idx]["elev_status"] = "falhou"
            continue
        k = max(1, math.ceil(len(profile) / 200))
        trails[idx]["elev"] = {
            "gain": round(gain), "loss": round(loss),
            "min": round(min(all_el)), "max": round(max(all_el)),
            "profile": [[round(d), round(e)] for d, e in profile[::k]],
        }
        trails[idx]["elev_status"] = "ok"


# ---------- 5. Dificuldade e tempo ----------
def classify(t: dict):
    sac = [s.strip() for s in t["sac_scale"].split(",") if s.strip() in SAC_ORDER]
    if sac:
        worst = max(SAC_ORDER.index(s) for s in sac)
        level = "facil" if worst == 0 else "moderada" if worst == 1 else "dificil"
        return level, "osm"
    gain = t["elev"]["gain"] if t["elev"] else None
    effort = t["length_km"] + (gain or 0) / 100
    level = "facil" if effort <= EASY_MAX else "moderada" if effort <= MODERATE_MAX else "dificil"
    return level, "calculado" if gain is not None else "calculado_sem_desnivel"


def est_time(t: dict):
    gain = t["elev"]["gain"] if t["elev"] else 0
    minutes = round((t["length_km"] / 5 + gain / 600) * 60)  # 5 km/h + 1 h por 600 m de subida
    return minutes, f"{minutes // 60}h{minutes % 60:02d}"


# ---------- 6. Saidas ----------
def write_sheet(wb, title, headers, rows, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in rows:
        ws.append(r)
    for i, h in enumerate(headers, 1):
        width = max([len(str(h))] + [len(str(r[i - 1])) for r in rows[:200]]) + 2
        ws.column_dimensions[get_column_letter(i)].width = min(width, 50)
    ws.freeze_panes = "A2"


def write_geojson(trails):
    feats = []
    for t in trails:
        el = t["elev"] or {}
        feats.append({
            "type": "Feature",
            "geometry": {"type": "MultiLineString",
                         "coordinates": [[[round(p[1], 6), round(p[0], 6)] for p in s]
                                         for s in t["segments"]]},
            "properties": {"name": t["name"], "km": t["length_km"],
                           "name_matches_regex": t["name_matches"],
                           "elevation_status": t["elev_status"],
                           "difficulty_calc": t["difficulty"], "gain_m": el.get("gain"),
                           "loss_m": el.get("loss"), "min_m": el.get("min"),
                           "max_m": el.get("max"), "time_calc": t["time"],
                           "profile": el.get("profile")},
        })
    OUT_GEOJSON.write_text(json.dumps({"type": "FeatureCollection", "features": feats},
                                      ensure_ascii=False), encoding="utf-8")


def main():
    data, meta = fetch()
    ways = parse_ways(data)
    print(f"{len(ways)} ways recebidos")

    groups = defaultdict(list)
    for w in ways:
        if w["key"]:
            groups[w["key"]].append(w)

    trails = []
    for g in groups.values():
        trails.extend(merge_group(g))
    trails.sort(key=lambda t: t["length_km"], reverse=True)

    if ELEVATION:
        add_elevation(trails)
    for t in trails:
        t["difficulty"], t["difficulty_source"] = classify(t)
        t["time_min"], t["time"] = est_time(t)

    total_pieces_node = sum(t["pieces_by_node"] for t in trails)
    with_elev = sum(t["elev"] is not None for t in trails)
    dist = defaultdict(int)
    for t in trails:
        dist[t["difficulty"]] += 1

    wb = Workbook()
    write_sheet(wb, "resumo", ["metrica", "valor"], [
        ["bbox (S,O,N,L)", BBOX],
        ["data da extracao (UTC)", meta.get("extracted_at", "")],
        ["timestamp da base OSM (Overpass)", meta.get("osm_base", "")],
        ["arquivo bruto usado", meta.get("cache_file", "")],
        ["highway consultados", ", ".join(meta.get("highways", []))],
        ["endpoints usados", ", ".join(sorted(set(meta.get("endpoints", {}).values())))],
        ["regex de nome (so marca, nao filtra)", NAME_REGEX or "(vazio)"],
        ["tolerancia GAP_M (m)", GAP_M],
        ["ways recebidos", len(ways)],
        ["nomes distintos", len(groups)],
        ["trilhas cujo nome bate no regex", sum(t["name_matches"] == "sim" for t in trails)],
        ["trilhas se juntar so por node compartilhado", total_pieces_node],
        ["trilhas se juntar por node + proximidade", len(trails)],
        ["juncoes extras feitas pela proximidade", total_pieces_node - len(trails)],
        ["trilhas que viraram 1 linha continua", sum(t["n_segments"] == 1 for t in trails)],
        ["trilhas com ramificacao/pedaco solto (>1 segmento)", sum(t["n_segments"] > 1 for t in trails)],
        ["trilhas circulares", sum(t["circular"] for t in trails)],
        ["trilhas >= 1 km", sum(t["length_km"] >= 1 for t in trails)],
        ["trilhas >= 3 km", sum(t["length_km"] >= 3 for t in trails)],
        ["elevacao: ok", with_elev],
        ["elevacao: pulada (abaixo de MIN_KM_ELEV)", sum(t["elev_status"] == "pulada" for t in trails)],
        ["elevacao: falhou", sum(t["elev_status"] == "falhou" for t in trails)],
        ["elevacao: desativada", sum(t["elev_status"] == "desativada" for t in trails)],
        ["dificuldade_calc: facil", dist["facil"]],
        ["dificuldade_calc: moderada", dist["moderada"]],
        ["dificuldade_calc: dificil", dist["dificil"]],
        ["dificuldade vinda do OSM (sac_scale)", sum(t["difficulty_source"] == "osm" for t in trails)],
    ], first=True)

    headers = ["nome", "nome_bate_regex", "n_ways", "pedacos_por_node", "segmentos", "km",
               "lat_centro", "lon_centro", "inicio_lat", "inicio_lon", "fim_lat", "fim_lon",
               "circular", "maior_lacuna_m", "dificuldade_calc", "fonte_dificuldade",
               "elevacao_status", "ganho_m", "perda_m", "min_m", "max_m", "tempo_est_calc",
               "ultima_edicao", "edicao_mais_antiga", "versao_max", "highway",
               *EXTRA_TAGS, "tags_json", "way_ids"]
    rows = []
    for t in trails:
        el = t["elev"] or {}
        rows.append([t["name"], t["name_matches"], t["n_ways"], t["pieces_by_node"], t["n_segments"],
                     t["length_km"], t["lat"], t["lon"], round(t["start"][0], 5), round(t["start"][1], 5),
                     round(t["end"][0], 5), round(t["end"][1], 5),
                     "sim" if t["circular"] else "nao", t["max_gap_m"],
                     t["difficulty"], t["difficulty_source"],
                     t["elev_status"], el.get("gain"), el.get("loss"), el.get("min"), el.get("max"),
                     t["time"], t["last_edit"], t["first_edit"], t["max_version"], t["highway"],
                     *[t[k] for k in EXTRA_TAGS], t["tags_json"], t["way_ids"]])
    write_sheet(wb, "trilhas", headers, rows)
    write_sheet(wb, "ways_brutos",
        ["way_id", "nome", "highway", "surface", "sac_scale", "metros", "n_nodes", "ultima_edicao"],
        [[w["id"], w["name"], w["highway"], w["tags"].get("surface", ""),
          w["tags"].get("sac_scale", ""), round(w["length_m"]), len(w["nodes"]),
          w["timestamp"]] for w in ways])
    wb.save(OUT)
    write_geojson(trails)
    print(f"\nOK -> {OUT} e {OUT_GEOJSON}")
    print(f"Trilhas: {len(trails)} (so por node: {total_pieces_node}); "
          f"elevacao em {with_elev}; dificuldade {dict(dist)}")


if __name__ == "__main__":
    main()