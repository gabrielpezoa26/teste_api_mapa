
# CLAUDE.md

## O que é este projeto

Spike para validar a ideia de um site de trilhas de São Paulo e entorno com dados do OpenStreetMap. Não é o app final (o app real vai usar SQL/PostGIS). O objetivo é responder:

1. Dá para extrair trilhas via API do Overpass, com foco em São Paulo e entorno?
2. Ways que se conectam no mundo real, mas aparecem quebrados na API, podem ser juntados?

A pesquisa que originou o spike está em `osm-trilhas.md` (decisões de arquitetura, tags, pipeline, fórmula de dificuldade).

## Estrutura

```
.
├── main.py              # todo o código (um arquivo só, na raiz)
├── Dockerfile           # python:3.12-slim, copia requirements.txt e main.py
├── docker-compose.yml   # serviço "app", monta .:/work, variáveis de ambiente
├── requirements.txt     # requests==2.32.3, openpyxl==3.1.5 (versões fixas)
├── Makefile             # atalhos (indentação com TAB)
├── osm-trilhas.md       # pesquisa de referência
└── data/                # criado pelo script; não versionar
    ├── overpass_raw.json     # cache da resposta do Overpass
    ├── elevation_cache.json  # cache da Open-Meteo, chave "lat,lon" com 5 casas
    ├── trilhas.xlsx          # saída principal
    └── trilhas.geojson       # saída para mapa (MultiLineString por trilha)
```

## Como rodar

Só Docker (roda em máquinas com versões diferentes). O daemon precisa estar ativo (`make start-docker`).

| Comando               | O que faz                                           |
| --------------------- | --------------------------------------------------- |
| `make`              | build + roda, usando o cache do Overpass se existir |
| `make refresh`      | ignora o cache e consulta o Overpass de novo        |
| `make gap GAP_M=60` | roda com outra tolerância de junção              |
| `make check-data`   | lista`data/`                                      |
| `make fclean`       | remove containers, imagem e`data/`                |

O volume `.:/work` monta a pasta do projeto, então `main.py` e `data/` ficam no disco do host. O script devolve a posse de `data/` ao dono da pasta do projeto ao terminar (`fix_owner`, via `atexit`), para não sobrar arquivo de root.

## Variáveis de ambiente (defaults no `main.py`)

| Variável       | Default                                 | Função                                                    |
| --------------- | --------------------------------------- | ----------------------------------------------------------- |
| `BBOX`        | `-24.05,-47.10,-23.00,-46.20`         | sul, oeste, norte, leste                                    |
| `GAP_M`       | `30`                                  | tolerância (m) para juntar pontas de ways                  |
| `NAME_REGEX`  | `trilha\|caminho\|pico\|cachoeira\|pedra` | filtro de nome (case-insensitive); vazio = todo way nomeado |
| `REFRESH`     | `0`                                   | `1` ignora o cache do Overpass                            |
| `ELEVATION`   | `1`                                   | `0` pula a Open-Meteo                                     |
| `STEP_M`      | `50`                                  | espaçamento da amostragem de elevação (m)                |
| `MIN_KM_ELEV` | `0.5`                                 | só calcula desnível de trilhas acima disso                |

O `docker-compose.yml` atual define só `BBOX`, `GAP_M`, `NAME_REGEX` e `REFRESH`; as demais usam o default do script.

## Fluxo do `main.py`

1. **`fetch()`**: POST no Overpass (3 endpoints públicos, 3 rodadas de retry). Query: `way[highway~"^(path|footway|track)$"][name~REGEX,i](BBOX); out meta geom;`. O `meta` traz `timestamp` e `version`; o `geom` traz coordenadas e a lista de `nodes`. Se o cache existir mas não tiver `timestamp`, refaz a consulta.
2. **`parse_ways()`**: um dict por way (tags completas, coords, nodes, comprimento por haversine).
3. **`merge_group()`**: agrupa por nome normalizado (sem acento, minúsculo). Dentro do nome, une ways em duas etapas com union-find: (a) node compartilhado, (b) pontas a até `GAP_M` metros. O resultado de cada cluster é uma trilha.
4. **`chain_ways()`**: encadeia os ways do cluster em linhas contínuas. Guloso: parte do way mais longo e cola o way cuja ponta esteja mais perto da cauda e depois da cabeça (inverte quando preciso), dentro de `GAP_M`. Sobras viram segmentos extras. `n_segments > 1` = ramificação ou pedaço que não encaixou.
5. **`add_elevation()`**: reamostra cada segmento (`resample`), busca a elevação na Open-Meteo em lotes de 100 (`fetch_elevations`, com cache), suaviza (média móvel de 5), aplica histerese de 4 m (`gain_loss`) e guarda `gain`, `loss`, `min`, `max` e um perfil de até 200 pontos.
6. **`classify()` / `est_time()`**: dificuldade pelo `sac_scale` quando existe (`osm`); senão por esforço `km + ganho/100` (≤6 fácil, ≤14 moderada, >14 difícil; fonte `calculado` ou `calculado_sem_desnivel`). Tempo estimado: 5 km/h mais 1 h a cada 600 m de subida, só ida.
7. **Saída**: `trilhas.xlsx` (abas `resumo`, `trilhas`, `ways_brutos`) e `trilhas.geojson`.

## Convenções

- Python 3.12, bibliotecas só `requests` e `openpyxl`. Sem shapely/geopandas de propósito (spike simples).
- Comentários, nomes de colunas do Excel e mensagens de log em português, sem acento nos identificadores e nas colunas.
- Valores em texto de dificuldade: `facil`, `moderada`, `dificil`.
- Não inventar métodos de API: usar só o que existe em `requests`, `openpyxl` e na Open-Meteo/Overpass.
- Ao alterar código, pedir ao Gabriel a versão atual do arquivo e editar em cima dela, em vez de mandar comandos de terminal para aplicar patch.

## Resultados observados (primeira execução real)

- 259 ways vindos do Overpass (igual à pesquisa), agrupados em 178 trilhas por node compartilhado e 175 depois da proximidade de 30 m. Os "ways quebrados" quase sempre já compartilham node; o gap resolve cerca de 2%.
- Elevação calculada em 88 de 175 trilhas (as outras são abaixo de 0,5 km, puladas de propósito).
- Dificuldade: 159 fácil, 13 moderada, 3 difícil (as fáceis incluem as curtas sem desnível).
- A Open-Meteo devolveu HTTP 429 em alguns lotes; o retry resolveu e o cache evita refazer.

## Limitações conhecidas

- Encadeamento guloso: bifurcações em T no meio de um way ficam como segmentos separados, e `inicio_*`/`fim_*` valem só para o segmento principal.
- Agrupamento só por nome; nomes genéricos em lugares distantes só se misturam se as pontas estiverem a menos de `GAP_M`.
- O filtro de nome perde trilhas sem "trilha", "caminho", "pico", "cachoeira" ou "pedra" no nome.
- DEM de 90 m (Copernicus GLO-90 via Open-Meteo): desnível é estimativa (±20% é esperado).
- Espera entre lotes da Open-Meteo é fixa (1 s) e a espera após 429 é `10 * tentativa` s.
- `highway=footway` em parque urbano gera ruído; ainda não há filtro de trilha urbana.

## Próximos passos possíveis

- Conferir no geojson.io as trilhas com mais de 1 segmento e as 3 junções feitas só por proximidade.
- Validar o `ganho_m` de 3 a 5 trilhas conhecidas contra referências públicas.
- Enriquecer com POIs, áreas protegidas e município (melhor no PostGIS), Wikidata, filtro de trilha urbana.
- Plano B de elevação: tiles Copernicus GLO-30 offline, sem limite de API.
- Migrar o pipeline para PostGIS (`ST_ClusterDBSCAN`, `ST_LineMerge`) e comparar com o resultado deste script.
