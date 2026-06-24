from flask import Flask, render_template, jsonify
from datetime import datetime, timedelta, timezone
import json
import os
import requests
import polars as pl
import csv
from io import StringIO

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
PERU_TZ = timezone(timedelta(hours=-5))
HORA_INICIO_ACTUALIZACION = 11

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
ruta_ranking = os.path.join(os.path.dirname(__file__), 'ranking_usuarios.json')
GOOGLE_SHEET_RANKING_CSV_URL = os.getenv(
    "GOOGLE_SHEET_RANKING_CSV_URL",
    "https://docs.google.com/spreadsheets/d/1A4fLL4bUPuu61HNzB8t3rm6o8RVUSJ89s2syxaQcacY/export?format=csv&gid=0"
)
def traducir_equipo(nombre):
    return TRADUCCION_PAISES.get(nombre, nombre)


def obtener_slot_actualizacion_peru(ahora_peru=None):
    ahora_peru = ahora_peru or datetime.now(PERU_TZ)
    if ahora_peru.hour < HORA_INICIO_ACTUALIZACION:
        return None
    return ahora_peru.replace(minute=0, second=0, microsecond=0).isoformat()


def _leer_ranking_desde_google_sheet():
    response = requests.get(GOOGLE_SHEET_RANKING_CSV_URL, timeout=20, verify=False)
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


def obtener_ranking_usuarios():
    try:
        return _leer_ranking_desde_google_sheet()
    except Exception as e:
        print(f"Error cargando ranking desde Google Sheet, usando JSON local: {e}")

    try:
        with open(ruta_ranking, 'r', encoding='utf-8') as f:
            registros_apuestas = json.load(f)
    except Exception as e:
        print(f"Error cargando ranking: {e}")
        registros_apuestas = []

    usuarios_consolidados = {}

    for reg in registros_apuestas:
        usuario = reg.get("CreatedBy")
        if not usuario:
            continue
            
        if usuario not in usuarios_consolidados:
            usuarios_consolidados[usuario] = {
                "nombre": usuario,
                "puntos": 0,
                "aciertos": 0
            }
        
        puntos = int(reg.get("PuntajeObtenido", 0))
        usuarios_consolidados[usuario]["puntos"] += puntos
        
        if puntos > 0:
            usuarios_consolidados[usuario]["aciertos"] += 1

    ranking_ordenado = sorted(
        list(usuarios_consolidados.values()),
        key=lambda x: (
            0 if x.get("puntos", 0) > 0 else 1,
            -x.get("puntos", 0),
            x.get("nombre", "").casefold()
        )
    )

    for idx, usuario in enumerate(ranking_ordenado, start=1):
        usuario["posicion"] = idx

    return {"ranking": ranking_ordenado}


def enviar_a_google_sheets(df_final):
    url_webhook = "https://script.google.com/macros/s/AKfycbz9RHYATwMmuL6jJkgOr59ucXZEB2cJ0RdVAKPk7qcMtq58M4ODZM-sRLK4DwMfbx8/exec"
    
    encabezados = [df_final.columns]
    filas = df_final.rows()
    data_final = encabezados + [list(fila) for fila in filas]

    try:
        response = requests.post(url_webhook, json=data_final, verify=False)
        print("Respuesta de Google:", response.status_code)
    except Exception as e:
        print("Error al enviar a Google Sheets:", e)


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
    df = pl.DataFrame(partidos)
    df_final = df.with_columns([
        pl.col("kickoff_utc").str.to_datetime("%Y-%m-%dT%H:%M:%S%.3fZ")
          .dt.strftime("%d/%m/%Y")
          .alias("kickoff_utc"),

        (pl.lit("'") + pl.col("hora_peru").str.slice(0, 5)).alias("hora_peru"),
        
        pl.when(pl.col("home_pen").is_not_null())
          .then(pl.lit("Penales"))
          .otherwise(pl.lit("Regular"))
          .alias("Tipo de Resultado"),
          
        pl.when(pl.col("status") == "completed")
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

        (pl.col("home_team") + " vs " + pl.col("away_team")).alias("Partido")
    ])
    ruta_excel = os.path.join(os.getcwd(), 'resultados_partidos.xlsx')
    df_final.write_excel(ruta_excel)
    enviar_a_google_sheets(df_final)

def calcular_posiciones_grupos(todos_los_partidos):
    estruct_grupos = {}

    for p in todos_los_partidos:
        g_name = p.get("group_name")
        if not g_name or p.get("round") != "group":
            continue
            
        if g_name not in estruct_grupos:
            estruct_grupos[g_name] = {}
            
        for equipo in [p["home_team"], p["away_team"]]:
            if equipo not in estruct_grupos[g_name]:
                estruct_grupos[g_name][equipo] = {"nombre": equipo, "pj": 0, "pts": 0, "gf": 0, "gc": 0, "dg": 0}

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

def obtener_datos_tablero():
    global cache_tablero, ultima_actualizacion_slot
    fecha_hoy_dt = datetime.now(PERU_TZ)
    slot_actualizacion = obtener_slot_actualizacion_peru(fecha_hoy_dt)
    debe_sincronizar = (
        not cache_tablero
        or (slot_actualizacion is not None and slot_actualizacion != ultima_actualizacion_slot)
    )
    
    if debe_sincronizar:
        print(f"Sincronizando marcadores con la API externa... slot Peru: {slot_actualizacion or 'carga inicial'}")
        
        url_api = "https://api.wc2026api.com/matches"
        token_real = "wc26_4TUutBnL1Qgocn3WrVSmmQ"
        headers = {"Authorization": f"Bearer {token_real}", "Accept": "application/json"}
        try:
            respuesta = requests.get(url_api, headers=headers, verify=False)
            resultados_en_vivo = respuesta.json() if respuesta.status_code == 200 else []
        except Exception as e:
            print(f"Error de conexión: {e}")
            resultados_en_vivo = []

        fecha_hoy_str = fecha_hoy_dt.strftime("%Y-%m-%d")
        
        diccionario_resultados = {r["id"]: r for r in resultados_en_vivo}
        
        for p in FIXTURE_ESTATICO:
            id_p = p.get("id")
            if id_p in diccionario_resultados:
                datos = diccionario_resultados[id_p]
                p.update({
                    "home_score": datos.get("home_score"),
                    "away_score": datos.get("away_score"),
                    "home_pen": datos.get("home_pen"),
                    "away_pen": datos.get("away_pen"),
                    "status": datos.get("status")
                })

        guardar_fixture_json(FIXTURE_ESTATICO)

        for p in FIXTURE_ESTATICO:
            p["home_team"] = traducir_equipo(p["home_team"])
            p["away_team"] = traducir_equipo(p["away_team"])

            fecha_utc_str = p.get("kickoff_utc")
            if fecha_utc_str:
                limpia = fecha_utc_str.split('.')[0] + 'Z' if '.' in fecha_utc_str else fecha_utc_str
                fecha_utc = datetime.strptime(limpia, "%Y-%m-%dT%H:%M:%SZ")
                fecha_peru_dt = fecha_utc - timedelta(hours=5)
                
                p["fecha_peru_str"] = f"{fecha_peru_dt.day} de {MESES[fecha_peru_dt.month]}"
                p["hora_peru"] = fecha_peru_dt.strftime("%H:%M")
                p["fecha_peru_key"] = fecha_peru_dt.strftime("%Y-%m-%d")

        partidos_ordenados = sorted(FIXTURE_ESTATICO, key=lambda x: x.get("kickoff_utc", ""))
        no_completados = [p for p in partidos_ordenados if p.get("status") != "completed"]

        if no_completados:
            principal = no_completados[0]
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

        fecha_base_str = principal.get("fecha_peru_key") if principal else fecha_hoy_str
        if not fecha_base_str and principal:
            fecha_base_str = principal.get("kickoff_utc", "")[:10]
        try:
            fecha_base_dt = datetime.strptime(fecha_base_str, "%Y-%m-%d")
        except Exception:
            fecha_base_dt = fecha_hoy_dt.replace(tzinfo=None)
        fecha_siguiente_str = (fecha_base_dt + timedelta(days=1)).strftime("%Y-%m-%d")
        fecha_anterior_str = (fecha_base_dt - timedelta(days=1)).strftime("%Y-%m-%d")

        ventana_resto = {fecha_base_str, fecha_siguiente_str}
        otros = [
            p for p in no_completados
            if p.get("id") not in ids_principales
            and p.get("fecha_peru_key", p.get("kickoff_utc", "")[:10]) in ventana_resto
        ]
        jugados = [
            p for p in partidos_ordenados
            if p.get("status") == "completed"
            and p.get("fecha_peru_key", p.get("kickoff_utc", "")[:10]) in {fecha_anterior_str, fecha_base_str}
        ]

        guardar_datos_excel(FIXTURE_ESTATICO)

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
        
        cache_tablero = {
            "principal": principal,
            "principales": principales,
            "otros": otros,
            "jugados": jugados,
            "grupos": tablas_grupos,
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
    response = jsonify(obtener_ranking_usuarios())
    response.headers["Cache-Control"] = "no-store"
    return response

if __name__ == '__main__':
    app.run(debug=True)
