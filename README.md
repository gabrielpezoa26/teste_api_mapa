
# Pipeline de Trilhas (OSM → Excel / GeoJSON / GPX / SQLite)

Coleta ways nomeados do OpenStreetMap, agrupa os pedaços de cada trilha, cruza a geometria com um DEM offline para calcular elevação e estima dificuldade e tempo. Objetivo: base **bruta**, nada é descartado na coleta.

## Fluxo

```
Overpass API ──► cache JSON ──► parse_ways ──► agrupar por nome ──► merge_group
                                                                        │
              Excel / GeoJSON / GPX / SQLite ◄── dificuldade/tempo ◄── elevação (DEM)
```

## Estrutura

```
main.py      # pipeline (coleta, merge, elevação, dificuldade, saídas)
utils.py     # geometria, Union-Find, encadeamento, suavização, planilhas
config.py    # todas as constantes e variáveis de ambiente
data/
├── utils/   # insumos: DEM (.tif/.vrt) e cache overpass_raw_*.json
├── xlsx/    # trilhas_AAAAMMDD.xlsx
├── geojson/ # trilhas_AAAAMMDD.geojson
├── sqlite/  # trilhas_AAAAMMDD.sqlite
└── gpx/gpx_AAAAMMDD/  # um .gpx por trilha
```

## Configuração (`config.py`)


| Variável    | Padrão                               | Função                                                      |
| -------------- | --------------------------------------- | --------------------------------------------------------------- |
| `BBOX`       | `-24.05,-47.10,-23.00,-46.20`         | Área da consulta (sul, oeste, norte, leste)                  |
| `HIGHWAYS`   | `path,footway,track`                  | Valores de`highway` consultados                               |
| `GAP_M`      | `30`                                  | Distância máxima (m) para considerar duas pontas conectadas |
| `NAME_REGEX` | `trilha|caminho|pico|cachoeira|pedra` | Só marca`nome_bate_regex`; **não filtra**                   |
| `REFRESH`    | `0`                                   | `1` ignora o cache e consulta o Overpass de novo              |
| `ELEVATION`  | `1`                                   | `0` desativa a etapa de elevação                            |
| `DEM_FILE`   | `data/utils/copernicus.vrt`           | Raster de elevação                                          |

Constantes fixas: `SMOOTH_RADIUS=2`, `HYSTERESIS_M=4.0`, `CIRCULAR_M=100`, `EASY_MAX=6`, `MODERATE_MAX=14`, velocidade 5 km/h e 600 m/h de subida.

## Etapas

### 1. Coleta (Overpass)

- Uma query por valor de `highway`: `way["highway"=X][name](BBOX); out meta geom;`
- Só ways **com nome**. `out meta geom` traz geometria (lat/lon de cada nó), IDs dos nós, versão e timestamp.
- Tenta 2 endpoints × 3 tentativas, com pausas e backoff. Se todos falharem, aborta.
- Resultados são deduplicados por `(type, id)` e salvos em `data/utils/overpass_raw_AAAAMMDD.json` com metadados.
- **Cache:** usa o arquivo mais recente se `BBOX` e `HIGHWAYS` forem iguais aos do cache e `REFRESH=0`.

### 2. Parse

- Cada way vira um dict: id, nome, chave normalizada, tags, nós, coordenadas `(lat, lon)` e comprimento em metros.
- Ways sem geometria ou sem nome normalizado são ignorados no agrupamento.

### 3. Agrupamento e merge

- Ways com a **mesma chave de nome** formam um grupo; cada grupo é dividido em **clusters** (uma trilha por cluster). Detalhes na seção de cálculos.
- Cada cluster gera um registro de trilha com: geometria (segmentos), comprimentos, início/fim, circular, centroide, tags consolidadas, IDs dos ways e datas de edição.
- Tags em `EXTRA_TAGS` viram colunas; valores distintos entre ways são unidos por `, `.

### 4. Elevação offline

- Detalhada na seção de cálculos. Se faltar `rasterio` ou o DEM, a trilha fica com `elev_status = "falhou (...)"` e o resto do pipeline segue.

### 5. Dificuldade e tempo

- Calculados por trilha (ver cálculos). Colunas `dificuldade_calc`, `fonte_dificuldade`, `tempo_est_calc`.

### 6. Saídas

- **Excel:** abas `resumo`, `trilhas`, `trilha_ways`, `ways_brutos`.
- **GeoJSON:** uma `Feature` `MultiLineString` por trilha (coordenadas com 6 casas), com propriedades resumidas.
- **GPX:** um arquivo por trilha (`TRL-<id>_AAAAMMDD.gpx`), um `<trkseg>` por segmento, com `<ele>` quando há elevação.
- **SQLite:** tabelas `trilhas` (inclui `geom_geojson`), `trilha_ways` (trilha ↔ way) e `ways`.

---

## Cálculos

### Distância (haversine)

- Distância entre dois pontos `(lat, lon)` sobre uma esfera de raio 6 371 000 m.
- `line_length` = soma da haversine entre pontos consecutivos.

### Normalização de nomes

- Remove acentos (NFKD → ASCII), minúsculas e espaços repetidos.
- "Trilha do Pico" e "trilha do  pico" caem no mesmo grupo.

### Agrupar ways em trilhas (`merge_group`)

Dois critérios de conexão, ambos via **Union-Find** (conjuntos disjuntos):

1. **Nó compartilhado:** se dois ways têm o mesmo ID de nó OSM, são unidos.
2. **Pontas próximas:** pontas (primeiro/último ponto) de ways distintos a ≤ `GAP_M` metros também são unidas.

Cada conjunto resultante é uma trilha. Ways de mesmo nome distantes entre si viram trilhas separadas.

### Encadeamento em segmentos (`chain_ways`)

- Ordena os ways do maior para o menor e parte do maior.
- Estende pela cauda e depois pela cabeça: procura a ponta mais próxima (≤ `GAP_M`) de qualquer way restante, inverte-o se preciso e anexa.
- Repete até não haver mais ponta próxima; o que sobra inicia um novo segmento.
- Resultado: lista de segmentos ordenada por comprimento. `max_gap_m` é o maior salto aceito.

### Comprimentos e metadados

- `km_total`: soma do comprimento de **todos** os ways do cluster.
- `km_principal`: comprimento do **maior segmento** encadeado. Todas as métricas derivadas (elevação, dificuldade, tempo) usam este.
- `circular`: distância entre início e fim do segmento principal `< 100 m`.
- `lat`/`lon` (centroide): média simples de todos os pontos.
- `trilha_id`: `TRL-<menor way_id do cluster>`.
- `ultima_edicao`: maior timestamp entre os ways.

### Cruzamento elevação × lat/lon

1. Para cada trilha, todos os pontos de **todos os segmentos** são concatenados numa lista, convertidos para `(lon, lat)` (ordem exigida pelo `rasterio`).
2. `src.sample(pontos)` consulta o DEM: transforma cada coordenada na posição de pixel do raster (via transformação afim do arquivo) e devolve o valor do pixel, em metros.
3. É **vizinho mais próximo**: sem interpolação, a altitude é a do pixel que contém o ponto (Copernicus GLO-30 ≈ 30 m de resolução).
4. A lista de altitudes é fatiada de volta por segmento, pelo comprimento de cada um (`elev_segments`), mantendo a correspondência ponto a ponto com a geometria. É isso que vai para o `<ele>` do GPX.
5. A amostragem ocorre **nos nós do OSM**; trechos com poucos nós têm menos pontos de altitude.
6. Pressupõe que o DEM esteja em coordenadas geográficas (lat/lon, WGS84), como o Copernicus.

### Ganho, perda e desnível (só segmento principal)

1. **Suavização:** média móvel com raio 2 (janela de até 5 pontos; nas bordas a janela é truncada).
2. **Histerese** (`gain_loss`): mantém uma altitude de referência. Quando a diferença para ela chega a ±4 m, soma ao ganho (ou perda) e a referência passa a ser o ponto atual. Oscilações menores que 4 m são ignoradas.
3. `desnivel_acumulado = max(ganho, perda)`.
4. `elev.min`/`elev.max` usam todos os pontos, mas não são exportados.

### Dificuldade

- **Se há `sac_scale` no OSM:** pega o pior valor entre os ways. `hiking` → fácil; `mountain_hiking` → moderada; acima disso → difícil. Fonte: `osm`.
- **Senão**, esforço = `km_principal + desnivel_acumulado / 100`:
  - `≤ 6` fácil · `≤ 14` moderada · `> 14` difícil
  - Fonte: `calculado` (ou `calculado_sem_desnivel` se não houve elevação).

### Tempo estimado

```
minutos = (km_principal / 5 + desnivel_acumulado / 600) × 60
```

- 5 km/h na horizontal + 600 m/h de subida, somados. Formato `HhMM`.

---

## Limitações conhecidas

- Se `src.sample` falhar, as altitudes viram `0.0` sem aviso: desnível 0 e status `ok`.
- Valores de *nodata* do DEM não são tratados.
- Sem interpolação no DEM; resolução de ~30 m limita a precisão.
- `merge_group` compara todas as pontas entre si (O(n²)): lento para áreas grandes com nomes repetidos.
- Dificuldade e tempo são heurísticos, não substituem avaliação real do terreno.
