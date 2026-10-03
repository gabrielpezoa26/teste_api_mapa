"""Overpass -> junta ways por nome -> Excel.

Duas formas de juntar ways do mesmo nome:
  1. por node compartilhado (conexao "de verdade" no OSM)
  2. por proximidade das pontas (GAP_M metros) - cobre ways que se
     encostam no mundo real mas nao compartilham node no mapa
O Excel mostra quantos pedacos existiam antes/depois de cada etapa.
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
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

BBOX = os.getenv("BBOX", "-24.05,-47.10,-23.00,-46.20")
GAP_M = float(os.getenv("GAP_M", "30"))
NAME_REGEX = os.getenv("NAME_REGEX", "trilha|caminho|pico|cachoeira|pedra")
REFRESH = os.getenv("REFRESH", "0") == "1"

DATA = Path("data")
RAW = DATA / "overpass_raw.json"
OUT = DATA / "trilhas.xlsx"


def fix_owner():
    """Devolve data/ ao dono da pasta do projeto (evita arquivos de root)."""
    try:
        st = os.stat(".")
        for f in [DATA, *DATA.glob("*")]:
            os.chown(f, st.st_uid, st.st_gid)
    except (OSError, AttributeError):
        pass


atexit.register(fix_owner)

ENDPOINTS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.private.coffee/api/interpreter",
]
HEADERS = {"User-Agent": "trilhas-sp-spike/0.1 (validacao de dados; contato: troque-aqui)"}


# ---------- 1. Overpass ----------
def build_query() -> str:
    name_filter = f'[name~"{NAME_REGEX}",i]' if NAME_REGEX else "[name]"
    return (
        "[out:json][timeout:180];\n"
        f'way[highway~"^(path|footway|track)$"]{name_filter}({BBOX});\n'
        "out body geom;"  # body traz a lista de nodes; geom traz as coordenadas
    )


def fetch() -> dict:
    DATA.mkdir(exist_ok=True)
    if RAW.exists() and not REFRESH:
        print(f"Usando cache {RAW} (REFRESH=1 para refazer)")
        return json.loads(RAW.read_text(encoding="utf-8"))

    query = build_query()
    print("Query:\n" + query)
    last_err = None
    for attempt in range(1, 4):
        for url in ENDPOINTS:
            try:
                print(f"[{attempt}/3] {url} ...")
                r = requests.post(url, data={"data": query}, headers=HEADERS, timeout=240)
                if r.status_code == 200:
                    data = r.json()
                    RAW.write_text(json.dumps(data), encoding="utf-8")
                    return data
                last_err = f"HTTP {r.status_code}"
            except (requests.RequestException, ValueError) as e:
                last_err = str(e)
            print(f"  falhou: {last_err}")
            time.sleep(5)
        time.sleep(15 * attempt)
    sys.exit(f"Overpass indisponivel: {last_err}. Tente de novo mais tarde.")


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


# ---------- 3. Juncao ----------
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
            "surface": tags.get("surface", ""),
            "sac_scale": tags.get("sac_scale", ""),
            "nodes": el.get("nodes", []),
            "coords": coords,
            "length_m": line_length(coords),
        })
    return ways


def merge_group(group: list[dict]) -> list[dict]:
    """Retorna lista de trilhas (clusters finais) para ways de um mesmo nome."""
    n = len(group)
    # etapa 1: node compartilhado
    uf_nodes = UF(n)
    owner = {}
    for i, w in enumerate(group):
        for nid in w["nodes"]:
            if nid in owner:
                uf_nodes.union(i, owner[nid])
            else:
                owner[nid] = i
    node_root = [uf_nodes.find(i) for i in range(n)]

    # etapa 2: proximidade das pontas (partindo dos clusters da etapa 1)
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
        result.append({
            "name": ws[0]["name"],
            "n_ways": len(ws),
            "pieces_by_node": len({node_root[i] for i in idxs}),
            "length_km": round(sum(w["length_m"] for w in ws) / 1000, 2),
            "lat": round(sum(p[0] for p in pts) / len(pts), 5),
            "lon": round(sum(p[1] for p in pts) / len(pts), 5),
            "highway": ", ".join(sorted({w["highway"] for w in ws if w["highway"]})),
            "surface": ", ".join(sorted({w["surface"] for w in ws if w["surface"]})),
            "sac_scale": ", ".join(sorted({w["sac_scale"] for w in ws if w["sac_scale"]})),
            "way_ids": ", ".join(str(w["id"]) for w in ws),
        })
    return result


# ---------- 4. Excel ----------
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
    return ws


def main():
    data = fetch()
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

    total_pieces_node = sum(t["pieces_by_node"] for t in trails)
    merged_by_gap = total_pieces_node - len(trails)

    wb = Workbook()
    write_sheet(wb, "resumo", ["metrica", "valor"], [
        ["bbox (S,O,N,L)", BBOX],
        ["filtro de nome", NAME_REGEX or "(todos nomeados)"],
        ["tolerancia GAP_M (m)", GAP_M],
        ["ways recebidos", len(ways)],
        ["nomes distintos", len(groups)],
        ["trilhas se juntar so por node compartilhado", total_pieces_node],
        ["trilhas se juntar por node + proximidade", len(trails)],
        ["juncoes extras feitas pela proximidade", merged_by_gap],
        ["trilhas >= 1 km", sum(t["length_km"] >= 1 for t in trails)],
        ["trilhas >= 3 km", sum(t["length_km"] >= 3 for t in trails)],
    ], first=True)
    write_sheet(wb, "trilhas",
        ["nome", "n_ways", "pedacos_por_node", "km", "lat_centro", "lon_centro",
         "highway", "surface", "sac_scale", "way_ids"],
        [[t["name"], t["n_ways"], t["pieces_by_node"], t["length_km"], t["lat"], t["lon"],
          t["highway"], t["surface"], t["sac_scale"], t["way_ids"]] for t in trails])
    write_sheet(wb, "ways_brutos",
        ["way_id", "nome", "highway", "surface", "sac_scale", "metros", "n_nodes"],
        [[w["id"], w["name"], w["highway"], w["surface"], w["sac_scale"],
          round(w["length_m"]), len(w["nodes"])] for w in ways])
    wb.save(OUT)
    print(f"\nOK -> {OUT}")
    print(f"Trilhas: {len(trails)} (so por node: {total_pieces_node}; "
          f"{merged_by_gap} juncoes extras por proximidade <= {GAP_M:.0f} m)")


if __name__ == "__main__":
    main()