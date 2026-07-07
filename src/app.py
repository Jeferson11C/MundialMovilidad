from flask import Flask, render_template, jsonify, request
from datetime import datetime, timedelta, timezone
import json
import os
import math
import threading
import copy
import requests
import polars as pl
import csv
from io import StringIO
from urllib.parse import quote
try:
    from .football_data_mappings import (
        GROUP_MATCH_NUMBER_BY_TEAMS,
        INTERNAL_ID_BY_MATCH_NUMBER,
        KNOCKOUT_START_MATCH_NUMBER,
        STADIUM_BY_MATCH_NUMBER,
        STAGE_TO_ROUND,
        STATUS_TO_INTERNAL,
        TEAM_CODE_ALIASES,
        TEAM_INFO_BY_CODE,
    )
except ImportError:
    from football_data_mappings import (
        GROUP_MATCH_NUMBER_BY_TEAMS,
        INTERNAL_ID_BY_MATCH_NUMBER,
        KNOCKOUT_START_MATCH_NUMBER,
        STADIUM_BY_MATCH_NUMBER,
        STAGE_TO_ROUND,
        STATUS_TO_INTERNAL,
        TEAM_CODE_ALIASES,
        TEAM_INFO_BY_CODE,
    )

app = Flask(__name__)

ruta_json = os.path.join(os.path.dirname(__file__), 'partidos_estaticos.json')

try:
    with open(ruta_json, 'r', encoding='utf-8') as f:
        FIXTURE_ESTATICO = json.load(f)
        print(f" Se cargaron {len(FIXTURE_ESTATICO)} partidos del archivo JSON")
except FileNotFoundError:
    print(f"No se encontró el archivo en la ruta: {ruta_json}")
    FIXTURE_ESTATICO = []

cache_tablero = {}
ultima_actualizacion_slot = None
sync_tablero_en_progreso = False
cache_lock = threading.Lock()
ranking_cache = {}
ranking_en_progreso = set()
RANKING_CACHE_TTL_SEGUNDOS = 300
PERU_TZ = timezone(timedelta(hours=-5))
INTERVALO_ACTUALIZACION_MINUTOS = 10

TRADUCCION_PAISES = {
    "Mexico": "México",
    "South Africa": "Sudáfrica",
    "Korea Republic": "Corea del Sur",
    "Czechia": "Chequia",
    "Canada": "Canadá",
    "Bosnia-Herzegovina": "Bosnia y Herzegovina",
    "USA": "Estados Unidos",
    "Paraguay": "Paraguay",
    "Qatar": "Catar",
    "Switzerland": "Suiza",
    "Brazil": "Brasil",
    "Morocco": "Marruecos",
    "Haiti": "Haití",
    "Scotland": "Escocia",
    "Australia": "Australia",
    "Turkey": "Turquía",
    "Germany": "Alemania",
    "Curaçao": "Curazao",
    "Netherlands": "Países Bajos",
    "Japan": "Japón",
    "Côte d'Ivoire": "Costa de Marfil",
    "Ecuador": "Ecuador",
    "Sweden": "Suecia",
    "Tunisia": "Túnez",
    "Spain": "España",
    "Cabo Verde": "Cabo Verde",
    "Belgium": "Bélgica",
    "Egypt": "Egipto",
    "Saudi Arabia": "Arabia Saudita",
    "Uruguay": "Uruguay",
    "IR Iran": "Irán",
    "New Zealand": "Nueva Zelanda",
    "France": "Francia",
    "Senegal": "Senegal",
    "Iraq": "Irak",
    "Norway": "Noruega",
    "Argentina": "Argentina",
    "Algeria": "Argelia",
    "Austria": "Austria",
    "Jordan": "Jordania",
    "Portugal": "Portugal",
    "Congo DR": "RD del Congo",
    "Uzbekistan": "Uzbekistán",
    "Colombia": "Colombia",
    "England": "Inglaterra",
    "Croatia": "Croacia",
    "Ghana": "Ghana",
    "Panama": "Panamá"
}

MESES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto", 
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
}
ruta_excel_partidos = os.path.join(os.path.dirname(__file__), 'resultados_partidos.xlsx')
GOOGLE_SHEET_RANKING_CSV_URL = os.getenv(
    "GOOGLE_SHEET_RANKING_CSV_URL",
    "https://docs.google.com/spreadsheets/d/1A4fLL4bUPuu61HNzB8t3rm6o8RVUSJ89s2syxaQcacY/export?format=csv&gid=0"
)
GOOGLE_SHEET_RANKING_FINAL_CSV_URL = os.getenv(
    "GOOGLE_SHEET_RANKING_FINAL_CSV_URL",
    "https://docs.google.com/spreadsheets/d/1A4fLL4bUPuu61HNzB8t3rm6o8RVUSJ89s2syxaQcacY/gviz/tq?tqx=out:csv&sheet="
    + quote("Hoja 2")
)
GOOGLE_SHEET_PARTIDOS_WEBHOOK_URL = os.getenv(
    "GOOGLE_SHEET_PARTIDOS_WEBHOOK_URL",
    "https://script.google.com/macros/s/AKfycbz9RHYATwMmuL6jJkgOr59ucXZEB2cJ0RdVAKPk7qcMtq58M4ODZM-sRLK4DwMfbx8/exec"
)
GOOGLE_SHEET_PARTIDOS_CSV_URL = os.getenv("GOOGLE_SHEET_PARTIDOS_CSV_URL", "")
FOOTBALL_DATA_MATCHES_URL = "https://api.football-data.org/v4/competitions/WC/matches"
FOOTBALL_DATA_TOKEN = os.getenv("FOOTBALL_DATA_TOKEN", "ef87be9ee90b4f0687ebf8524cb92252")
EXCEL_COLUMNAS_PARTIDOS = [
    "id", "match_number", "round", "group_name",
    "home_team_id", "home_team", "home_team_code", "home_team_flag",
    "away_team_id", "away_team", "away_team_code", "away_team_flag",
    "stadium_id", "stadium", "stadium_city", "stadium_country",
    "kickoff_utc", "home_score", "away_score", "home_pen", "away_pen",
    "status", "fecha_peru_str", "hora_peru", "fecha_peru_key",
    "Tipo de Resultado", "Ganador", "Partido",
]
EQUIPOS_CLASIFICADOS_BASE = {
    "Alemania", "Paraguay", "Francia", "Suecia", "Sudáfrica", "Canadá",
    "Países Bajos", "Marruecos", "Portugal", "Croacia", "España", "Austria",
    "Estados Unidos", "Bosnia y Herzegovina", "Bélgica", "Senegal",
    "Brasil", "Japón", "Costa de Marfil", "Noruega", "México", "Ecuador",
    "Inglaterra", "RD del Congo", "Argentina", "Cabo Verde", "Australia",
    "Egipto", "Suiza", "Argelia", "Colombia", "Ghana",
    "Germany", "South Africa", "Canada", "Netherlands", "Spain", "USA",
    "Bosnia-Herzegovina", "Belgium", "Japan", "Côte d'Ivoire", "Mexico",
    "England", "Congo DR", "Egypt", "Switzerland", "Algeria",
}
CODIGOS_CLASIFICADOS_BASE = {
    "GER", "PAR", "FRA", "SWE", "RSA", "CAN", "NED", "MAR",
    "POR", "CRO", "ESP", "AUT", "USA", "BIH", "BEL", "SEN",
    "BRA", "JPN", "CIV", "NOR", "MEX", "ECU", "ENG", "COD",
    "ARG", "CPV", "AUS", "EGY", "SUI", "ALG", "COL", "GHA",
}


def traducir_equipo(nombre):
    return TRADUCCION_PAISES.get(nombre, nombre)


def obtener_slot_actualizacion_peru(ahora_peru=None):
    ahora_peru = ahora_peru or datetime.now(PERU_TZ)
    minuto_slot = (ahora_peru.minute // INTERVALO_ACTUALIZACION_MINUTOS) * INTERVALO_ACTUALIZACION_MINUTOS
    return ahora_peru.replace(minute=minuto_slot, second=0, microsecond=0).isoformat()


def convertir_utc_a_peru(fecha_utc_str):
    fecha_utc = datetime.fromisoformat(fecha_utc_str.replace("Z", "+00:00"))
    if fecha_utc.tzinfo is None:
        fecha_utc = fecha_utc.replace(tzinfo=timezone.utc)
    return fecha_utc.astimezone(PERU_TZ)


def formatear_fecha_key_sheet(fecha_peru_dt):
    return f"{fecha_peru_dt.day}/{fecha_peru_dt.month:02d}/{fecha_peru_dt.year}"


def _leer_ranking_desde_google_sheet(url_csv):
    response = requests.get(url_csv, timeout=20, verify=False)
    response.raise_for_status()

    contenido = response.content.decode("utf-8-sig")
    reader = csv.DictReader(StringIO(contenido))
    ranking = []

    for fila in reader:
        nombre = (fila.get("usuario") or "").strip()
        if not nombre:
            continue

        try:
            aciertos = int(float((fila.get("aciertos") or "0").strip()))
        except Exception:
            aciertos = 0

        try:
            puntos = int(float((fila.get("puntos") or "0").strip()))
        except Exception:
            puntos = 0

        ranking.append({
            "nombre": nombre,
            "aciertos": aciertos,
            "puntos": puntos
        })

    ranking.sort(key=lambda x: (
        0 if x.get("puntos", 0) > 0 else 1,
        -x.get("puntos", 0),
        x.get("nombre", "").casefold()
    ))
    for idx, usuario in enumerate(ranking, start=1):
        usuario["posicion"] = idx

    return {"ranking": ranking}


def actualizar_ranking_en_background(fase, url_ranking):
    global ranking_cache, ranking_en_progreso
    try:
        datos = _leer_ranking_desde_google_sheet(url_ranking)
        with cache_lock:
            ranking_cache[fase] = {
                "datos": datos,
                "actualizado": datetime.now(PERU_TZ),
            }
    except Exception as e:
        print(f"Error cargando ranking desde Google Sheet: {e}")
    finally:
        with cache_lock:
            ranking_en_progreso.discard(fase)


def obtener_ranking_usuarios(fase="grupos"):
    url_ranking = (
        GOOGLE_SHEET_RANKING_FINAL_CSV_URL
        if fase == "final"
        else GOOGLE_SHEET_RANKING_CSV_URL
    )
    ahora = datetime.now(PERU_TZ)
    with cache_lock:
        cache = ranking_cache.get(fase)
        en_progreso = fase in ranking_en_progreso
        if cache and (ahora - cache["actualizado"]).total_seconds() < RANKING_CACHE_TTL_SEGUNDOS:
            return {**cache["datos"], "actualizando": False}
        if not en_progreso:
            ranking_en_progreso.add(fase)
            threading.Thread(
                target=actualizar_ranking_en_background,
                args=(fase, url_ranking),
                daemon=True,
            ).start()

    if cache:
        return {**cache["datos"], "actualizando": True}
    return {"ranking": [], "actualizando": True}


def enviar_a_google_sheets(df_final):
    encabezados = [df_final.columns]
    filas = df_final.rows()
    data_final = encabezados + [list(fila) for fila in filas]

    try:
        response = requests.post(GOOGLE_SHEET_PARTIDOS_WEBHOOK_URL, json=data_final, verify=False)
        print("Respuesta de Google:", response.status_code)
    except Exception as e:
        print("Error al enviar a Google Sheets:", e)


def limpiar_valor_sheet(valor):
    if valor is None:
        return None
    if isinstance(valor, float) and math.isnan(valor):
        return None
    if isinstance(valor, str):
        valor = valor.strip()
        return valor if valor else None
    return valor


def convertir_entero_sheet(valor):
    valor = limpiar_valor_sheet(valor)
    if valor is None:
        return None
    try:
        return int(float(str(valor).replace(",", ".")))
    except (TypeError, ValueError):
        return None


def aplicar_estadio_por_match_number(partido):
    match_number = convertir_entero_sheet(partido.get("match_number"))
    estadio = STADIUM_BY_MATCH_NUMBER.get(match_number)
    if not estadio:
        return partido

    partido["match_number"] = match_number
    partido["stadium_id"] = partido.get("stadium_id") or estadio[0]
    partido["stadium"] = partido.get("stadium") or estadio[1]
    partido["stadium_city"] = partido.get("stadium_city") or estadio[2]
    partido["stadium_country"] = partido.get("stadium_country") or estadio[3]
    return partido


def obtener_filas_sheet_desde_respuesta(respuesta):
    texto = respuesta.text.strip()
    if not texto:
        return []

    content_type = respuesta.headers.get("Content-Type", "")
    if "json" in content_type.lower() or texto.startswith("[") or texto.startswith("{"):
        data = respuesta.json()
        if isinstance(data, dict):
            data = data.get("rows") or data.get("data") or data.get("values") or []
        if not data:
            return []
        if all(isinstance(fila, dict) for fila in data):
            return data
        if all(isinstance(fila, list) for fila in data):
            encabezados = [str(columna).strip() for columna in data[0]]
            return [
                dict(zip(encabezados, fila))
                for fila in data[1:]
                if any(limpiar_valor_sheet(valor) is not None for valor in fila)
            ]
        return []

    lector = csv.DictReader(StringIO(respuesta.content.decode("utf-8-sig")))
    return [
        fila
        for fila in lector
        if any(limpiar_valor_sheet(valor) is not None for valor in fila.values())
    ]


def normalizar_partido_desde_sheet(fila):
    partido_id = convertir_entero_sheet(fila.get("id"))
    if partido_id is None:
        return None

    partido = {columna: limpiar_valor_sheet(fila.get(columna)) for columna in EXCEL_COLUMNAS_PARTIDOS}
    for columna in [
        "id", "match_number", "home_team_id", "away_team_id", "stadium_id",
        "home_score", "away_score", "home_pen", "away_pen",
    ]:
        partido[columna] = convertir_entero_sheet(partido.get(columna))

    kickoff_utc = partido.get("kickoff_utc")
    if not isinstance(kickoff_utc, str) or "T" not in kickoff_utc:
        partido["kickoff_utc"] = None

    partido["home_team_flag"] = None
    partido["away_team_flag"] = None
    aplicar_estadio_por_match_number(partido)

    if partido.get("home_team") and partido.get("away_team") and not partido.get("Partido"):
        partido["Partido"] = f"{partido['home_team']} vs {partido['away_team']}"

    return partido


def obtener_partidos_desde_google_sheet():
    urls = []
    if GOOGLE_SHEET_PARTIDOS_CSV_URL:
        urls.append(("CSV", GOOGLE_SHEET_PARTIDOS_CSV_URL))
    if GOOGLE_SHEET_PARTIDOS_WEBHOOK_URL:
        urls.append(("webhook", GOOGLE_SHEET_PARTIDOS_WEBHOOK_URL))

    for origen, url in urls:
        try:
            print(f"Consultando fallback Google Sheet de partidos ({origen})...")
            respuesta = requests.get(url, timeout=20, verify=False)
            print(f"Google Sheet partidos ({origen}) status code: {respuesta.status_code}")
            respuesta.raise_for_status()
            filas = obtener_filas_sheet_desde_respuesta(respuesta)
            partidos = [
                partido
                for partido in (normalizar_partido_desde_sheet(fila) for fila in filas)
                if partido is not None
            ]
            print(f"Partidos leidos desde Google Sheet: {len(partidos)}")
            if partidos:
                return partidos
        except Exception as e:
            print(f"Error leyendo Google Sheet de partidos ({origen}): {e}")

    return []


def obtener_partidos_para_sincronizar(fecha_hoy_dt):
    partidos = obtener_partidos_desde_api()
    if partidos:
        return partidos

    print("La API no devolvio partidos. Se conserva la data local y no se actualiza el Sheet.")
    return []


def normalizar_codigo_equipo(codigo):
    if not codigo:
        return None
    codigo = codigo.strip().upper()
    return TEAM_CODE_ALIASES.get(codigo, codigo)


def obtener_info_equipo(codigo, errores_mapeo):
    codigo_normalizado = normalizar_codigo_equipo(codigo)
    if not codigo_normalizado:
        return {"id": None, "nombre": None, "codigo": None}

    info = TEAM_INFO_BY_CODE.get(codigo_normalizado)
    if not info:
        errores_mapeo.append(f"Equipo sin mapeo interno: {codigo_normalizado}")
        return {"id": None, "nombre": None, "codigo": codigo_normalizado}

    return {"id": info["id"], "nombre": info["nombre"], "codigo": codigo_normalizado}


def obtener_match_number_grupo(home_code, away_code, errores_mapeo):
    match_number = GROUP_MATCH_NUMBER_BY_TEAMS.get((home_code, away_code))
    if match_number is None:
        errores_mapeo.append(f"Cruce de grupos sin match_number: {home_code} vs {away_code}")
    return match_number


def construir_indices_eliminatoria(matches):
    indices = {}
    for ronda, inicio in KNOCKOUT_START_MATCH_NUMBER.items():
        stage = next((k for k, v in STAGE_TO_ROUND.items() if v == ronda), None)
        partidos_ronda = sorted(
            [m for m in matches if STAGE_TO_ROUND.get(m.get("stage")) == ronda],
            key=lambda m: m.get("utcDate") or ""
        )
        for offset, match in enumerate(partidos_ronda):
            indices[match.get("id")] = inicio + offset
    return indices


def obtener_tipo_resultado(match):
    duration = ((match.get("score") or {}).get("duration") or "").upper()
    if duration == "PENALTY_SHOOTOUT":
        return "Penales"
    if duration == "EXTRA_TIME":
        return "Tiempo extra"
    return "Regular"


def obtener_ganador(match, home_name, away_name, status):
    winner = ((match.get("score") or {}).get("winner") or "").upper()
    if status == "scheduled":
        return "Por jugar"
    if status == "in_progress":
        return "En vivo"
    if winner == "HOME_TEAM":
        return home_name or "Por definir"
    if winner == "AWAY_TEAM":
        return away_name or "Por definir"
    if winner == "DRAW":
        return "Empate"
    return "Por jugar"


def obtener_group_name(match):
    group = match.get("group")
    if not group:
        return None
    return group.replace("GROUP_", "")[-1]


def calcular_marcador_base(full_time, penalties):
    home_score = full_time.get("home")
    away_score = full_time.get("away")
    home_pen = penalties.get("home")
    away_pen = penalties.get("away")
    tiene_penales = home_pen is not None and away_pen is not None

    if tiene_penales and home_score is not None and away_score is not None:
        home_score = home_score - home_pen
        away_score = away_score - away_pen

    return home_score, away_score, home_pen, away_pen


def inferir_estado_por_marcador(status, kickoff_utc, home_score, away_score, ahora_peru=None):
    if status in {"completed", "in_progress"}:
        return status
    if home_score is None or away_score is None or not kickoff_utc:
        return status

    try:
        ahora_peru = ahora_peru or datetime.now(PERU_TZ)
        inicio_peru = convertir_utc_a_peru(kickoff_utc)
    except Exception:
        return "in_progress"

    if inicio_peru <= ahora_peru <= inicio_peru + timedelta(hours=4):
        return "in_progress"
    if ahora_peru > inicio_peru + timedelta(hours=4):
        return "completed"
    return status


def obtener_ganador_desde_partido(partido):
    status = partido.get("status")
    if status == "scheduled":
        return "Por jugar"
    if status in {"live", "in_progress"}:
        return "En vivo"

    home_team = partido.get("home_team") or "Por definir"
    away_team = partido.get("away_team") or "Por definir"
    home_score = partido.get("home_score")
    away_score = partido.get("away_score")
    home_pen = partido.get("home_pen")
    away_pen = partido.get("away_pen")

    if status == "completed":
        if home_pen is not None and away_pen is not None:
            return home_team if home_pen > away_pen else away_team
        if home_score is not None and away_score is not None:
            if home_score > away_score:
                return home_team
            if away_score > home_score:
                return away_team
            return "Empate"
    return "Por jugar"


def actualizar_campos_derivados_partido(partido):
    partido["Partido"] = (
        f"{partido.get('home_team') or 'Por definir'} vs "
        f"{partido.get('away_team') or 'Por definir'}"
    )
    partido["Ganador"] = obtener_ganador_desde_partido(partido)
    partido["marcador_visual"] = construir_marcador_visual(
        partido.get("home_team") or "Por definir",
        partido.get("away_team") or "Por definir",
        partido.get("home_score"),
        partido.get("away_score"),
        partido.get("home_pen"),
        partido.get("away_pen"),
    )
    return partido


def construir_marcador_visual(home_team, away_team, home_score, away_score, home_pen, away_pen):
    if home_score is None or away_score is None:
        return None
    if home_pen is not None and away_pen is not None:
        return f"{home_team} {home_score}({home_pen}) - ({away_pen}){away_score} {away_team}"
    return f"{home_team} {home_score} - {away_score} {away_team}"


def transformar_partido_football_data(match, match_number, errores_mapeo):
    home = match.get("homeTeam") or {}
    away = match.get("awayTeam") or {}
    home_info = obtener_info_equipo(home.get("tla"), errores_mapeo)
    away_info = obtener_info_equipo(away.get("tla"), errores_mapeo)
    score = match.get("score") or {}
    full_time = score.get("fullTime") or {}
    penalties = score.get("penalties") or {}
    home_score, away_score, home_pen, away_pen = calcular_marcador_base(full_time, penalties)
    round_name = STAGE_TO_ROUND.get(match.get("stage"))
    status = STATUS_TO_INTERNAL.get(match.get("status"), "scheduled")
    status = inferir_estado_por_marcador(
        status,
        match.get("utcDate"),
        home_score,
        away_score,
    )
    internal_id = INTERNAL_ID_BY_MATCH_NUMBER.get(match_number)
    stadium_id, stadium, stadium_city, stadium_country = STADIUM_BY_MATCH_NUMBER.get(
        match_number,
        (None, None, None, None)
    )

    if match_number is None:
        errores_mapeo.append(f"Partido sin match_number: external_id={match.get('id')}")
    elif internal_id is None:
        errores_mapeo.append(f"match_number sin id interno: {match_number}")

    return {
        "id": internal_id,
        "match_number": match_number,
        "round": round_name,
        "group_name": obtener_group_name(match) if round_name == "group" else None,
        "home_team_id": home_info["id"],
        "home_team": home_info["nombre"],
        "home_team_code": home_info["codigo"],
        "home_team_flag": home.get("crest"),
        "away_team_id": away_info["id"],
        "away_team": away_info["nombre"],
        "away_team_code": away_info["codigo"],
        "away_team_flag": away.get("crest"),
        "stadium_id": stadium_id,
        "stadium": stadium,
        "stadium_city": stadium_city,
        "stadium_country": stadium_country,
        "kickoff_utc": match.get("utcDate"),
        "home_score": home_score,
        "away_score": away_score,
        "home_pen": home_pen,
        "away_pen": away_pen,
        "status": status,
        "Tipo de Resultado": obtener_tipo_resultado(match),
        "Ganador": obtener_ganador(match, home_info["nombre"], away_info["nombre"], status),
        "Partido": f"{home_info['nombre'] or 'Por definir'} vs {away_info['nombre'] or 'Por definir'}",
        "marcador_visual": construir_marcador_visual(
            home_info["nombre"] or "Por definir",
            away_info["nombre"] or "Por definir",
            home_score,
            away_score,
            home_pen,
            away_pen,
        ),
    }


def obtener_partidos_desde_api():
    headers = {"X-Auth-Token": FOOTBALL_DATA_TOKEN, "Accept": "application/json"}
    print(f"Consultando football-data.org: {FOOTBALL_DATA_MATCHES_URL}")

    try:
        respuesta = requests.get(
            FOOTBALL_DATA_MATCHES_URL,
            headers=headers,
            timeout=20,
            verify=False
        )
        print(f"football-data.org status code: {respuesta.status_code}")

        if respuesta.status_code == 401:
            print("Error football-data.org: token invalido")
            return []
        if respuesta.status_code == 403:
            print("Error football-data.org: sin permisos")
            return []
        if respuesta.status_code == 429:
            print("Error football-data.org: rate limit alcanzado")
            return []
        if respuesta.status_code >= 500:
            print("Error football-data.org: error del servidor")
            return []
        respuesta.raise_for_status()
    except Exception as e:
        print(f"Error consultando football-data.org: {e}")
        return []

    payload = respuesta.json()
    matches = payload.get("matches", [])
    print(f"Partidos recibidos football-data.org: {len(matches)}")

    errores_mapeo = []
    indices_eliminatoria = construir_indices_eliminatoria(matches)
    partidos = []

    for match in matches:
        round_name = STAGE_TO_ROUND.get(match.get("stage"))
        if not round_name:
            errores_mapeo.append(f"Stage sin mapeo: {match.get('stage')}")
            continue

        home_code = normalizar_codigo_equipo((match.get("homeTeam") or {}).get("tla"))
        away_code = normalizar_codigo_equipo((match.get("awayTeam") or {}).get("tla"))
        if round_name == "group":
            match_number = obtener_match_number_grupo(home_code, away_code, errores_mapeo)
        else:
            match_number = indices_eliminatoria.get(match.get("id"))

        partido = transformar_partido_football_data(match, match_number, errores_mapeo)
        if partido.get("id") is not None:
            partidos.append(partido)

    estados = {}
    for partido in partidos:
        estado = partido.get("status") or "sin_estado"
        estados[estado] = estados.get(estado, 0) + 1

    print(f"Filas generadas/actualizadas: {len(partidos)}")
    print(f"Partidos por estado: {estados}")
    for error in errores_mapeo:
        print(f"Error de mapeo: {error}")

    return partidos


def guardar_fixture_json(partidos):
    traducciones_invertidas = {v: k for k, v in TRADUCCION_PAISES.items()}
    campos_derivados = {"fecha_peru_str", "hora_peru", "fecha_peru_key"}
    partidos_json = []

    for partido in partidos:
        partido_limpio = {
            clave: valor
            for clave, valor in partido.items()
            if clave not in campos_derivados
        }
        partido_limpio["home_team"] = traducciones_invertidas.get(
            partido_limpio.get("home_team"),
            partido_limpio.get("home_team")
        )
        partido_limpio["away_team"] = traducciones_invertidas.get(
            partido_limpio.get("away_team"),
            partido_limpio.get("away_team")
        )
        partidos_json.append(partido_limpio)

    try:
        with open(ruta_json, 'w', encoding='utf-8') as f:
            json.dump(partidos_json, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Error guardando fixture JSON: {e}")


def guardar_datos_excel(partidos):
    filas = []
    for partido in partidos:
        aplicar_estadio_por_match_number(partido)
        fila = {columna: partido.get(columna) for columna in EXCEL_COLUMNAS_PARTIDOS}
        aplicar_estadio_por_match_number(fila)
        fila["home_team_flag"] = None
        fila["away_team_flag"] = None
        fila["home_team"] = traducir_equipo(fila.get("home_team"))
        fila["away_team"] = traducir_equipo(fila.get("away_team"))

        fecha_utc_str = fila.get("kickoff_utc")
        if fecha_utc_str:
            try:
                fecha_peru_dt = convertir_utc_a_peru(fecha_utc_str)
                fila["kickoff_utc"] = fecha_peru_dt.strftime("%d/%m/%Y")
                fila["fecha_peru_str"] = f"{fecha_peru_dt.day} de {MESES[fecha_peru_dt.month]}"
                fila["hora_peru"] = fecha_peru_dt.strftime("%H:%M")
                fila["fecha_peru_key"] = formatear_fecha_key_sheet(fecha_peru_dt)
            except Exception:
                fila["kickoff_utc"] = str(fecha_utc_str)[:10]
                if fila.get("hora_peru"):
                    fila["hora_peru"] = str(fila.get("hora_peru"))[:5]

        filas.append(fila)

    df = pl.DataFrame(filas)
    df_final = df.with_columns([
        pl.col("hora_peru").fill_null("").str.slice(0, 5).alias("hora_peru"),
        
        pl.when(pl.col("Tipo de Resultado") == "Tiempo extra")
          .then(pl.lit("Tiempo extra"))
          .when(pl.col("home_pen").is_not_null())
          .then(pl.lit("Penales"))
          .otherwise(pl.lit("Regular"))
          .alias("Tipo de Resultado"),
          
        pl.when(pl.col("status") == "in_progress")
            .then(pl.lit("En vivo"))
          .when(pl.col("status") == "completed")
            .then(
                pl.when(pl.col("home_pen").is_not_null())
                .then(
                    pl.when(pl.col("home_pen") > pl.col("away_pen")).then(pl.col("home_team"))
                    .otherwise(pl.col("away_team"))
                )
                .otherwise(
                    pl.when(pl.col("home_score") > pl.col("away_score")).then(pl.col("home_team"))
                    .when(pl.col("away_score") > pl.col("home_score")).then(pl.col("away_team"))
                    .otherwise(pl.lit("Empate"))
                )
          ).otherwise(pl.lit("Por jugar"))
          .alias("Ganador"),

        (
            pl.col("home_team").fill_null("Por definir")
            + " vs "
            + pl.col("away_team").fill_null("Por definir")
        ).alias("Partido")
    ]).select(EXCEL_COLUMNAS_PARTIDOS)
    df_final.write_excel(ruta_excel_partidos)
    enviar_a_google_sheets(df_final)

def calcular_posiciones_grupos(todos_los_partidos):
    estruct_grupos = {}

    for p in todos_los_partidos:
        g_name = p.get("group_name")
        if not g_name or p.get("round") != "group":
            continue
            
        if g_name not in estruct_grupos:
            estruct_grupos[g_name] = {}
            
        for lado in ["home", "away"]:
            equipo = p.get(f"{lado}_team")
            codigo = p.get(f"{lado}_team_code")
            if not equipo:
                continue
            if equipo not in estruct_grupos[g_name]:
                estruct_grupos[g_name][equipo] = {
                    "nombre": equipo,
                    "codigo": codigo,
                    "pj": 0,
                    "pts": 0,
                    "gf": 0,
                    "gc": 0,
                    "dg": 0,
                }

    for p in todos_los_partidos:
        g_name = p.get("group_name")
        if not g_name or p.get("round") != "group" or p.get("status") != "completed":
            continue
            
        e1, e2 = p["home_team"], p["away_team"]
        g1, g2 = int(p["home_score"]), int(p["away_score"])
        
        if e1 in estruct_grupos[g_name]:
            estruct_grupos[g_name][e1]["pj"] += 1
            estruct_grupos[g_name][e1]["gf"] += g1
            estruct_grupos[g_name][e1]["gc"] += g2
        if e2 in estruct_grupos[g_name]:
            estruct_grupos[g_name][e2]["pj"] += 1
            estruct_grupos[g_name][e2]["gf"] += g2
            estruct_grupos[g_name][e2]["gc"] += g1
        
        if g1 > g2 and e1 in estruct_grupos[g_name]:
            estruct_grupos[g_name][e1]["pts"] += 3
        elif g1 < g2 and e2 in estruct_grupos[g_name]:
            estruct_grupos[g_name][e2]["pts"] += 3
        elif g1 == g2:
            if e1 in estruct_grupos[g_name]:
                estruct_grupos[g_name][e1]["pts"] += 1
            if e2 in estruct_grupos[g_name]:
                estruct_grupos[g_name][e2]["pts"] += 1

    grupos_ordenados = []
    for g_name, equipos_dict in estruct_grupos.items():
        lista_equipos = list(equipos_dict.values())
        
        for eq in lista_equipos:
            eq["dg"] = eq["gf"] - eq["gc"]
        
        lista_equipos.sort(key=lambda x: (x["pts"], x["dg"]), reverse=True)
        grupos_ordenados.append({"nombre": f"Grupo {g_name}", "equipos": lista_equipos})
        
    grupos_ordenados.sort(key=lambda x: x["nombre"])
    return grupos_ordenados


def construir_cache_tablero(fecha_hoy_dt):
    fecha_hoy_str = fecha_hoy_dt.strftime("%Y-%m-%d")
    fixture_cache = copy.deepcopy(FIXTURE_ESTATICO)

    for p in fixture_cache:
        aplicar_estadio_por_match_number(p)
        if p.get("home_team"):
            p["home_team"] = traducir_equipo(p["home_team"])
        if p.get("away_team"):
            p["away_team"] = traducir_equipo(p["away_team"])

        fecha_utc_str = p.get("kickoff_utc")
        p["status"] = inferir_estado_por_marcador(
            p.get("status"),
            fecha_utc_str,
            p.get("home_score"),
            p.get("away_score"),
            fecha_hoy_dt,
        )
        if fecha_utc_str:
            fecha_peru_dt = convertir_utc_a_peru(fecha_utc_str)
            p["fecha_peru_str"] = f"{fecha_peru_dt.day} de {MESES[fecha_peru_dt.month]}"
            p["hora_peru"] = fecha_peru_dt.strftime("%H:%M")
            p["fecha_peru_key"] = fecha_peru_dt.strftime("%Y-%m-%d")
        actualizar_campos_derivados_partido(p)

    partidos_ordenados = sorted(fixture_cache, key=lambda x: x.get("kickoff_utc", ""))
    partidos_grupo = [p for p in partidos_ordenados if p.get("round") == "group"]
    partidos_finales = [p for p in partidos_ordenados if p.get("round") != "group"]
    no_completados = [p for p in partidos_ordenados if p.get("status") != "completed"]

    partidos_en_vivo = [p for p in no_completados if p.get("status") in {"live", "in_progress"}]
    partidos_programados = [
        p for p in no_completados
        if p.get("fecha_peru_key", p.get("kickoff_utc", "")[:10]) >= fecha_hoy_str
    ]
    if partidos_en_vivo:
        principal = partidos_en_vivo[0]
    elif partidos_programados:
        principal = partidos_programados[0]
    else:
        principal = partidos_ordenados[-1] if partidos_ordenados else None

    principal_fecha_key = principal.get("fecha_peru_key", principal.get("kickoff_utc", "")[:10]) if principal else None
    principal_hora = principal.get("hora_peru") if principal else None
    principales = [
        p for p in no_completados
        if p.get("fecha_peru_key", p.get("kickoff_utc", "")[:10]) == principal_fecha_key
        and p.get("hora_peru") == principal_hora
    ] if principal else []
    ids_principales = {p.get("id") for p in principales}

    fecha_anterior_str = (fecha_hoy_dt - timedelta(days=1)).strftime("%Y-%m-%d")
    fecha_siguiente_str = (fecha_hoy_dt + timedelta(days=1)).strftime("%Y-%m-%d")
    ventana_proximos = {fecha_hoy_str, fecha_siguiente_str}
    ventana_jugados = {fecha_anterior_str, fecha_hoy_str}

    otros = [
        p for p in no_completados
        if p.get("id") not in ids_principales
        and p.get("fecha_peru_key", p.get("kickoff_utc", "")[:10]) in ventana_proximos
    ]
    jugados = [
        p for p in partidos_ordenados
        if p.get("status") == "completed"
        and p.get("fecha_peru_key", p.get("kickoff_utc", "")[:10]) in ventana_jugados
    ]

    es_fase_grupos = fecha_hoy_str < "2026-06-28"
    partidos_eliminatoria = []
    if not es_fase_grupos:
        eliminatorias = [p for p in fixture_cache if p.get("round") != "group"]
        ronda_activa = "R32"
        for p in eliminatorias:
            if p.get("kickoff_utc", "")[:10] >= fecha_hoy_str:
                ronda_activa = p.get("round")
                break
        if fecha_hoy_str > "2026-07-19":
            ronda_activa = "final"
        partidos_eliminatoria = [p for p in fixture_cache if p.get("round") == ronda_activa]

    tablas_grupos = calcular_posiciones_grupos(fixture_cache)
    equipos_clasificados = {
        equipo
        for p in partidos_finales
        if p.get("round") == "R32"
        for equipo in (p.get("home_team"), p.get("away_team"))
        if equipo
    }
    equipos_clasificados.update(EQUIPOS_CLASIFICADOS_BASE)
    for grupo in tablas_grupos:
        for equipo in grupo.get("equipos", []):
            equipo["clasificado"] = (
                equipo.get("codigo") in CODIGOS_CLASIFICADOS_BASE
                or equipo.get("nombre") in equipos_clasificados
            )

    return {
        "principal": principal,
        "principales": principales,
        "otros": otros,
        "jugados": jugados,
        "grupos": tablas_grupos,
        "partidos_grupo": partidos_grupo,
        "partidos_finales": partidos_finales,
        "es_fase_grupos": es_fase_grupos,
        "partidos_eliminatoria": partidos_eliminatoria,
    }


def sincronizar_tablero_en_background(slot_actualizacion):
    global cache_tablero, ultima_actualizacion_slot, sync_tablero_en_progreso
    try:
        fecha_hoy_dt = datetime.now(PERU_TZ)
        print(f"Sincronizando marcadores en background... slot Peru: {slot_actualizacion}")
        resultados_en_vivo = obtener_partidos_para_sincronizar(fecha_hoy_dt)

        diccionario_resultados = {
            r.get("id"): r
            for r in resultados_en_vivo
            if r.get("id") is not None
        }

        if diccionario_resultados:
            for p in FIXTURE_ESTATICO:
                id_p = p.get("id")
                if id_p in diccionario_resultados:
                    datos = diccionario_resultados[id_p]
                    p.update({
                        "match_number": datos.get("match_number", p.get("match_number")) or p.get("match_number"),
                        "round": datos.get("round", p.get("round")) or p.get("round"),
                        "group_name": datos.get("group_name", p.get("group_name")) or p.get("group_name"),
                        "home_team_id": datos.get("home_team_id", p.get("home_team_id")) or p.get("home_team_id"),
                        "home_team": datos.get("home_team", p.get("home_team")) or p.get("home_team"),
                        "home_team_code": datos.get("home_team_code", p.get("home_team_code")) or p.get("home_team_code"),
                        "home_team_flag": datos.get("home_team_flag", p.get("home_team_flag")) or p.get("home_team_flag"),
                        "away_team_id": datos.get("away_team_id", p.get("away_team_id")) or p.get("away_team_id"),
                        "away_team": datos.get("away_team", p.get("away_team")) or p.get("away_team"),
                        "away_team_code": datos.get("away_team_code", p.get("away_team_code")) or p.get("away_team_code"),
                        "away_team_flag": datos.get("away_team_flag", p.get("away_team_flag")) or p.get("away_team_flag"),
                        "stadium_id": datos.get("stadium_id", p.get("stadium_id")) or p.get("stadium_id"),
                        "stadium": datos.get("stadium", p.get("stadium")) or p.get("stadium"),
                        "stadium_city": datos.get("stadium_city", p.get("stadium_city")) or p.get("stadium_city"),
                        "stadium_country": datos.get("stadium_country", p.get("stadium_country")) or p.get("stadium_country"),
                        "kickoff_utc": datos.get("kickoff_utc", p.get("kickoff_utc")) or p.get("kickoff_utc"),
                        "home_score": datos.get("home_score"),
                        "away_score": datos.get("away_score"),
                        "home_pen": datos.get("home_pen"),
                        "away_pen": datos.get("away_pen"),
                        "status": datos.get("status", p.get("status")) or p.get("status"),
                        "Tipo de Resultado": datos.get("Tipo de Resultado", p.get("Tipo de Resultado")) or p.get("Tipo de Resultado"),
                        "Ganador": datos.get("Ganador", p.get("Ganador")) or p.get("Ganador"),
                        "Partido": datos.get("Partido", p.get("Partido")) or p.get("Partido"),
                        "marcador_visual": datos.get("marcador_visual", p.get("marcador_visual")) or p.get("marcador_visual"),
                    })

            for p in FIXTURE_ESTATICO:
                aplicar_estadio_por_match_number(p)
            guardar_fixture_json(FIXTURE_ESTATICO)
            guardar_datos_excel(FIXTURE_ESTATICO)
        else:
            print("Sheet no actualizado porque la API no devolvio datos.")

        nuevo_cache = construir_cache_tablero(datetime.now(PERU_TZ))
        with cache_lock:
            cache_tablero = nuevo_cache
            ultima_actualizacion_slot = slot_actualizacion
    finally:
        with cache_lock:
            sync_tablero_en_progreso = False


def obtener_datos_tablero():
    global cache_tablero, ultima_actualizacion_slot, sync_tablero_en_progreso
    fecha_hoy_dt = datetime.now(PERU_TZ)
    slot_actualizacion = obtener_slot_actualizacion_peru(fecha_hoy_dt)
    with cache_lock:
        if not cache_tablero:
            cache_tablero = construir_cache_tablero(fecha_hoy_dt)

        debe_refrescar = slot_actualizacion != ultima_actualizacion_slot
        if debe_refrescar and not sync_tablero_en_progreso:
            sync_tablero_en_progreso = True
            threading.Thread(
                target=sincronizar_tablero_en_background,
                args=(slot_actualizacion,),
                daemon=True,
            ).start()

        return {**cache_tablero, "actualizando": sync_tablero_en_progreso}

    debe_sincronizar = (
        not cache_tablero
        or (slot_actualizacion is not None and slot_actualizacion != ultima_actualizacion_slot)
    )
    
    if debe_sincronizar:
        print(f"Sincronizando marcadores... slot Peru: {slot_actualizacion or 'carga inicial'}")
        
        resultados_en_vivo = obtener_partidos_para_sincronizar(fecha_hoy_dt)

        fecha_hoy_str = fecha_hoy_dt.strftime("%Y-%m-%d")
        
        diccionario_resultados = {
            r.get("id"): r
            for r in resultados_en_vivo
            if r.get("id") is not None
        }
        
        for p in FIXTURE_ESTATICO:
            id_p = p.get("id")
            if id_p in diccionario_resultados:
                datos = diccionario_resultados[id_p]
                p.update({
                    "match_number": datos.get("match_number", p.get("match_number")) or p.get("match_number"),
                    "round": datos.get("round", p.get("round")) or p.get("round"),
                    "group_name": datos.get("group_name", p.get("group_name")) or p.get("group_name"),
                    "home_team_id": datos.get("home_team_id", p.get("home_team_id")) or p.get("home_team_id"),
                    "home_team": datos.get("home_team", p.get("home_team")) or p.get("home_team"),
                    "home_team_code": datos.get("home_team_code", p.get("home_team_code")) or p.get("home_team_code"),
                    "home_team_flag": datos.get("home_team_flag", p.get("home_team_flag")) or p.get("home_team_flag"),
                    "away_team_id": datos.get("away_team_id", p.get("away_team_id")) or p.get("away_team_id"),
                    "away_team": datos.get("away_team", p.get("away_team")) or p.get("away_team"),
                    "away_team_code": datos.get("away_team_code", p.get("away_team_code")) or p.get("away_team_code"),
                    "away_team_flag": datos.get("away_team_flag", p.get("away_team_flag")) or p.get("away_team_flag"),
                    "stadium_id": datos.get("stadium_id", p.get("stadium_id")) or p.get("stadium_id"),
                    "stadium": datos.get("stadium", p.get("stadium")) or p.get("stadium"),
                    "stadium_city": datos.get("stadium_city", p.get("stadium_city")) or p.get("stadium_city"),
                    "stadium_country": datos.get("stadium_country", p.get("stadium_country")) or p.get("stadium_country"),
                    "kickoff_utc": datos.get("kickoff_utc", p.get("kickoff_utc")) or p.get("kickoff_utc"),
                    "home_score": datos.get("home_score"),
                    "away_score": datos.get("away_score"),
                    "home_pen": datos.get("home_pen"),
                    "away_pen": datos.get("away_pen"),
                    "status": datos.get("status", p.get("status")) or p.get("status"),
                    "Tipo de Resultado": datos.get("Tipo de Resultado", p.get("Tipo de Resultado")) or p.get("Tipo de Resultado"),
                    "Ganador": datos.get("Ganador", p.get("Ganador")) or p.get("Ganador"),
                    "Partido": datos.get("Partido", p.get("Partido")) or p.get("Partido"),
                    "marcador_visual": datos.get("marcador_visual", p.get("marcador_visual")) or p.get("marcador_visual")
                })

        for p in FIXTURE_ESTATICO:
            aplicar_estadio_por_match_number(p)

        guardar_fixture_json(FIXTURE_ESTATICO)

        for p in FIXTURE_ESTATICO:
            aplicar_estadio_por_match_number(p)
            p["home_team"] = traducir_equipo(p["home_team"])
            p["away_team"] = traducir_equipo(p["away_team"])

            fecha_utc_str = p.get("kickoff_utc")
            if fecha_utc_str:
                fecha_peru_dt = convertir_utc_a_peru(fecha_utc_str)
                
                p["fecha_peru_str"] = f"{fecha_peru_dt.day} de {MESES[fecha_peru_dt.month]}"
                p["hora_peru"] = fecha_peru_dt.strftime("%H:%M")
                p["fecha_peru_key"] = fecha_peru_dt.strftime("%Y-%m-%d")

        partidos_ordenados = sorted(FIXTURE_ESTATICO, key=lambda x: x.get("kickoff_utc", ""))
        partidos_grupo = [p for p in partidos_ordenados if p.get("round") == "group"]
        partidos_finales = [p for p in partidos_ordenados if p.get("round") != "group"]
        no_completados = [p for p in partidos_ordenados if p.get("status") != "completed"]

        partidos_en_vivo = [p for p in no_completados if p.get("status") in {"live", "in_progress"}]
        partidos_programados = [
            p for p in no_completados
            if p.get("fecha_peru_key", p.get("kickoff_utc", "")[:10]) >= fecha_hoy_str
        ]
        if partidos_en_vivo:
            principal = partidos_en_vivo[0]
        elif partidos_programados:
            principal = partidos_programados[0]
        else:
            principal = partidos_ordenados[-1] if partidos_ordenados else None

        principal_fecha_key = principal.get("fecha_peru_key", principal.get("kickoff_utc", "")[:10]) if principal else None
        principal_hora = principal.get("hora_peru") if principal else None
        principales = [
            p for p in no_completados
            if p.get("fecha_peru_key", p.get("kickoff_utc", "")[:10]) == principal_fecha_key
            and p.get("hora_peru") == principal_hora
        ] if principal else []
        ids_principales = {p.get("id") for p in principales}

        fecha_anterior_str = (fecha_hoy_dt - timedelta(days=1)).strftime("%Y-%m-%d")
        fecha_siguiente_str = (fecha_hoy_dt + timedelta(days=1)).strftime("%Y-%m-%d")
        ventana_proximos = {fecha_hoy_str, fecha_siguiente_str}
        ventana_jugados = {fecha_anterior_str, fecha_hoy_str}

        otros = [
            p for p in no_completados
            if p.get("id") not in ids_principales
            and p.get("fecha_peru_key", p.get("kickoff_utc", "")[:10]) in ventana_proximos
        ]
        jugados = [
            p for p in partidos_ordenados
            if p.get("status") == "completed"
            and p.get("fecha_peru_key", p.get("kickoff_utc", "")[:10]) in ventana_jugados
        ]

        if resultados_en_vivo:
            guardar_datos_excel(FIXTURE_ESTATICO)
        else:
            print("Sheet no actualizado porque la API no devolvio datos.")

        es_fase_grupos = True
        partidos_eliminatoria = []
        
        if fecha_hoy_str >= "2026-06-28":
            es_fase_grupos = False
            eliminatorias = [p for p in FIXTURE_ESTATICO if p.get("round") != "group"]
            ronda_activa = "R32"
            for p in eliminatorias:
                if p.get("kickoff_utc", "")[:10] >= fecha_hoy_str:
                    ronda_activa = p.get("round")
                    break
            if fecha_hoy_str > "2026-07-19": ronda_activa = "final"
            partidos_eliminatoria = [p for p in FIXTURE_ESTATICO if p.get("round") == ronda_activa]
        
        tablas_grupos = calcular_posiciones_grupos(FIXTURE_ESTATICO)
        equipos_clasificados = {
            equipo
            for p in partidos_finales
            if p.get("round") == "R32"
            for equipo in (p.get("home_team"), p.get("away_team"))
            if equipo
        }
        equipos_clasificados.update(EQUIPOS_CLASIFICADOS_BASE)
        for grupo in tablas_grupos:
            for equipo in grupo.get("equipos", []):
                equipo["clasificado"] = (
                    equipo.get("codigo") in CODIGOS_CLASIFICADOS_BASE
                    or equipo.get("nombre") in equipos_clasificados
                )
        
        cache_tablero = {
            "principal": principal,
            "principales": principales,
            "otros": otros,
            "jugados": jugados,
            "grupos": tablas_grupos,
            "partidos_grupo": partidos_grupo,
            "partidos_finales": partidos_finales,
            "es_fase_grupos": es_fase_grupos,
            "partidos_eliminatoria": partidos_eliminatoria 
        }
        ultima_actualizacion_slot = slot_actualizacion
        
    return cache_tablero

@app.route('/')
def home():
    return render_template('index.html')


@app.route('/ranking')
def ranking():
    return render_template('ranking.html')


@app.route('/api/datos_tablero')
def api_tablero():
    return jsonify(obtener_datos_tablero())


@app.route('/api/ranking')
def api_ranking():
    fase = request.args.get("fase", "grupos")
    if fase not in {"grupos", "final"}:
        fase = "grupos"
    response = jsonify(obtener_ranking_usuarios(fase))
    response.headers["Cache-Control"] = "no-store"
    return response

if __name__ == '__main__':
    app.run(debug=True)
