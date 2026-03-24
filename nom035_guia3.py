"""
NOM-035-STPS-2018 | RFRANYUTTI, CONCIENCIA VERDE Y LABORAL S.C.
================================================================
MÓDULO 0 + 3  —  Panel Operativo + Cuestionario Virtual
GUÍA III      —  Identificación y Análisis de los Factores de
                 Riesgo Psicosocial y Evaluación del Entorno
                 Organizacional en los Centros de Trabajo
(Para centros de trabajo con más de 50 trabajadores)
"""

import streamlit as st
import pandas as pd
import os, re, time, sys, io, textwrap, threading
from datetime import datetime

# ── Modo empleado ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="NOM-035 · Guía III", page_icon="🏭",
                   layout="centered", initial_sidebar_state="collapsed")

_params        = st.query_params
_MODO_EMPLEADO = "cliente" in _params
_CLIENTE_URL   = str(_params.get("cliente", "")).upper() if _MODO_EMPLEADO else ""

# ── Pantalla de carga ─────────────────────────────────────────────────────────
if "app_loaded_g3" not in st.session_state:
    st.session_state.app_loaded_g3 = False
if not st.session_state.app_loaded_g3:
    _ph = st.empty()
    _ph.markdown("""
    <div class="loading-overlay">
        <div class="loading-logo" style="font-family:Montserrat,sans-serif;font-size:1.1rem;
             font-weight:700;color:#4b694e;letter-spacing:.08em;">RFRANYUTTI</div>
        <div style="font-size:.72rem;color:#888;font-family:Montserrat,sans-serif;
             font-weight:500;margin-top:.3rem;letter-spacing:.1em;">
            CONCIENCIA VERDE Y LABORAL S.C.</div>
        <div class="loading-ring"></div>
        <div class="loading-txt">NOM-035-STPS-2018 · Guía III · Cargando...</div>
    </div>
    """, unsafe_allow_html=True)
    import time as _t; _t.sleep(1.2)
    _ph.empty()
    st.session_state.app_loaded_g3 = True

# ── Resolución de rutas ───────────────────────────────────────────────────────
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def _asset(rel):
    nombre = os.path.basename(rel)
    for c in [
        os.path.join(_BASE_DIR, rel),
        os.path.join(_BASE_DIR, nombre),
        os.path.join(_BASE_DIR, "assets", "logos", nombre),
        os.path.join(_BASE_DIR, "assets", nombre),
        os.path.join(os.getcwd(), rel),
        os.path.join(os.getcwd(), nombre),
        os.path.join(os.getcwd(), "assets", "logos", nombre),
        rel,
    ]:
        if c and os.path.exists(c):
            return c
    return os.path.join(_BASE_DIR, rel)

# ── Clientes ──────────────────────────────────────────────────────────────────
CLIENTES = {
    "FRUCO": {
        "razon": "FRUTAS CONCENTRADAS, S.A.P.I. DE C.V.",
        "logo":  _asset("assets/logos/fruco.png"),
        "opciones": ["FRUTAS CONCENTRADAS, S.A.P.I. DE C.V."],
    },
    "QUALTIA": {
        "razon": "QUALTIA ALIMENTOS Y OPERACIONES, S. DE R.L. DE C.V.",
        "logo":  _asset("assets/logos/qualtia.png"),
        "opciones": [
            "QUALTIA ALIMENTOS Y OPERACIONES, S. DE R.L. DE C.V.",
            "QUALTIA ALIMENTOS OPERACIONES, S. DE R.L. DE C.V. (CEDIS Y SERVICIOS AUXILIARES)",
        ],
    },
    "DIABLOS": {
        "razon": "CENTRO DEPORTIVO ALFREDO HARP HELÚ, S.A. DE C.V.",
        "logo":  _asset("assets/logos/Diablos.png"),
        "opciones": ["CENTRO DEPORTIVO ALFREDO HARP HELÚ, S.A. DE C.V."],
    },
}
LOGO_RF = _asset("assets/logos/rfranyutti.gif")

def excel_path(cliente_key: str, razon_social: str = "") -> str:
    if cliente_key == "QUALTIA" and "CEDIS" in razon_social.upper():
        return "data/g3_resultados_QUALTIA_CEDIS.csv"
    return f"data/g3_resultados_{cliente_key.upper()}.csv"

# ── Catálogos ──────────────────────────────────────────────────────────────────
SEL         = "— Selecciona —"
OPC_RESP    = ["Siempre", "Casi siempre", "Algunas veces", "Casi nunca", "Nunca"]
OPC_SEXO    = [SEL, "Femenino", "Masculino"]
OPC_EDAD    = [SEL,"15 - 19","20 - 24","25 - 29","30 - 34","35 - 39","40 - 44",
               "45 - 49","50 - 54","55 - 59","60 - 64","65 - 69","70 o más"]
OPC_ECIVIL  = [SEL,"Soltero","Casado","Unión libre","Divorciado","Viudo"]
OPC_ESTUD   = [SEL,"Sin formación","Primaria","Secundaria","Preparatoria o Bachillerato",
               "Técnico Superior","Licenciatura","Maestría","Doctorado"]
OPC_PUESTO  = [SEL,"Operativo","Supervisor","Profesional o técnico","Gerente"]
OPC_CONTRAT = [SEL,"Por obra o proyecto","Tiempo indeterminado",
               "Por tiempo determinado (temporal)","Honorarios"]
OPC_PERSONAL= [SEL,"Sindicalizado","Confianza","Ninguno"]
OPC_JORNADA = [SEL,"Fijo diurno (entre las 6:00 y 20:00 hrs)",
               "Fijo nocturno (entre las 20:00 y 6:00 hrs)",
               "Fijo mixto (combinación de nocturno y diurno)"]
OPC_TPUESTO = [SEL,"Menos de 6 meses","Entre 6 meses y 1 año","Entre 1 a 4 años",
               "Entre 5 a 9 años","Entre 10 a 14 años","Entre 15 a 19 años",
               "Entre 20 a 24 años","25 años o más"]
OPC_EXP     = [SEL,"Menos de 6 meses","Entre 6 meses y 1 año","Entre 1 a 4 años",
               "Entre 5 a 9 años","Entre 10 a 14 años","Entre 15 a 19 años",
               "Entre 20 a 24 años","25 años o más"]

# ── Tabla 5 NOM-035 Guía III — Escala de calificación ────────────────────────
# INVERSOS (Siempre=0 ... Nunca=4):
# 1,4,23,24,25,26,27,28,30,31,32,33,34,35,36,37,38,39,40,41,
# 42,43,44,45,46,47,48,49,50,51,52,53,55,56,57
ITEMS_INVERSOS_G3 = {
    1, 4, 23, 24, 25, 26, 27, 28, 30, 31, 32, 33, 34, 35, 36,
    37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51,
    52, 53, 55, 56, 57
}
# DIRECTOS (Siempre=4 ... Nunca=0): todos los demás
ITEMS_DIRECTOS_G3 = set(range(1, 73)) - ITEMS_INVERSOS_G3

ESCALA_D_G3 = {"Siempre": 4, "Casi siempre": 3, "Algunas veces": 2, "Casi nunca": 1, "Nunca": 0}
ESCALA_I_G3 = {"Siempre": 0, "Casi siempre": 1, "Algunas veces": 2, "Casi nunca": 3, "Nunca": 4}

# ── Tabla 6 — Dominios y categorías ──────────────────────────────────────────
# dom_id: 0=Ambiente, 1=Carga, 2=Control, 3=Jornada, 4=Interferencia,
#         5=Liderazgo, 6=Relaciones, 7=Violencia, 8=Reconocimiento, 9=Pertenencia
NOMBRES_DOM_G3 = {
    0: "Condiciones en el ambiente de trabajo",
    1: "Carga de trabajo",
    2: "Falta de control sobre el trabajo",
    3: "Jornada de trabajo",
    4: "Interferencia en la relación trabajo-familia",
    5: "Liderazgo",
    6: "Relaciones en el trabajo",
    7: "Violencia laboral",
    8: "Reconocimiento del desempeño",
    9: "Insuficiente sentido de pertenencia e inestabilidad",
}

# Tabla 6 NOM-035 — Puntos de corte por dominio (Cdom)
TABLA_DOM_G3 = {
    0: (4,  8,  10, 13),   # Ambiente          Cdom<5,<9,<11,<14
    1: (14, 20, 26, 36),   # Carga             Cdom<15,<21,<27,<37
    2: (10, 15, 20, 24),   # Control           Cdom<11,<16,<21,<25
    3: (0,  1,  3,  5),    # Jornada           Cdom<1,<2,<4,<6
    4: (3,  5,  7,  9),    # Interferencia     Cdom<4,<6,<8,<10
    5: (8,  11, 15, 19),   # Liderazgo         Cdom<9,<12,<16,<20
    6: (9,  12, 16, 20),   # Relaciones        Cdom<10,<13,<17,<21
    7: (6,  9,  12, 15),   # Violencia         Cdom<7,<10,<13,<16
    8: (5,  9,  13, 17),   # Reconocimiento    Cdom<6,<10,<14,<18
    9: (3,  5,  7,  9),    # Pertenencia       Cdom<4,<6,<8,<10
}

# Tabla — Puntos de corte por categoría (Ccat)
TABLA_CAT_G3 = {
    1: (4,  8,  10, 13),   # Ambiente          Ccat<5,<9,<11,<14
    2: (14, 29, 44, 59),   # Factores          Ccat<15,<30,<45,<60
    3: (4,  6,  9,  12),   # Organización      Ccat<5,<7,<10,<13
    4: (13, 28, 41, 57),   # Liderazgo/relac   Ccat<14,<29,<42,<58
    5: (9,  13, 17, 22),   # Entorno org.      Ccat<10,<14,<18,<23
}

# Tabla 7 — Calificación final (Cfinal)
TABLA5_G3 = [
    (0,   49,  "NULO",     "#4B694E", "#D6E4D8"),   # Cfinal < 50
    (50,  74,  "BAJO",     "#69A2D8", "#EBF3FB"),   # 50 < Cfinal < 75
    (75,  98,  "MEDIO",    "#C8A600", "#FFF8DC"),   # 75 < Cfinal < 99
    (99,  139, "ALTO",     "#E07820", "#FDEBD0"),   # 99 < Cfinal < 140
    (140, 999, "MUY ALTO", "#A20000", "#FDDEDE"),   # Cfinal > 140
]

# ── Preguntas oficiales Guía III (72 ítems) ───────────────────────────────────
PREGUNTAS_G3 = [
    # SECCIÓN 1: Condiciones ambientales (ítems 1-5)
    {"id":1,  "cat":1, "dom":0, "sec":"Condiciones ambientales",
     "cat_nombre":"Ambiente de trabajo",
     "dom_nombre":"Condiciones en el ambiente de trabajo",
     "texto":"El espacio donde trabajo me permite realizar mis actividades de manera segura e higiénica."},
    {"id":2,  "cat":1, "dom":0, "sec":"Condiciones ambientales",
     "cat_nombre":"Ambiente de trabajo",
     "dom_nombre":"Condiciones en el ambiente de trabajo",
     "texto":"Mi trabajo me exige hacer mucho esfuerzo físico."},
    {"id":3,  "cat":1, "dom":0, "sec":"Condiciones ambientales",
     "cat_nombre":"Ambiente de trabajo",
     "dom_nombre":"Condiciones en el ambiente de trabajo",
     "texto":"Me preocupa sufrir un accidente en mi trabajo."},
    {"id":4,  "cat":1, "dom":0, "sec":"Condiciones ambientales",
     "cat_nombre":"Ambiente de trabajo",
     "dom_nombre":"Condiciones en el ambiente de trabajo",
     "texto":"Considero que en mi trabajo se aplican las normas de seguridad y salud en el trabajo."},
    {"id":5,  "cat":1, "dom":0, "sec":"Condiciones ambientales",
     "cat_nombre":"Ambiente de trabajo",
     "dom_nombre":"Condiciones en el ambiente de trabajo",
     "texto":"Considero que las actividades que realizo son peligrosas."},

    # SECCIÓN 2: Cantidad y ritmo de trabajo (ítems 6-8)
    {"id":6,  "cat":2, "dom":1, "sec":"Cantidad y ritmo de trabajo",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Por la cantidad de trabajo que tengo debo quedarme tiempo adicional a mi turno."},
    {"id":7,  "cat":2, "dom":1, "sec":"Cantidad y ritmo de trabajo",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Por la cantidad de trabajo que tengo debo trabajar sin parar."},
    {"id":8,  "cat":2, "dom":1, "sec":"Cantidad y ritmo de trabajo",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Considero que es necesario mantener un ritmo de trabajo acelerado."},

    # SECCIÓN 3: Esfuerzo mental (ítems 9-12)
    {"id":9,  "cat":2, "dom":1, "sec":"Esfuerzo mental",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Mi trabajo exige que esté muy concentrado."},
    {"id":10, "cat":2, "dom":1, "sec":"Esfuerzo mental",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Mi trabajo requiere que memorice mucha información."},
    {"id":11, "cat":2, "dom":1, "sec":"Esfuerzo mental",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"En mi trabajo tengo que tomar decisiones difíciles muy rápido."},
    {"id":12, "cat":2, "dom":1, "sec":"Esfuerzo mental",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Mi trabajo exige que atienda varios asuntos al mismo tiempo."},

    # SECCIÓN 4: Actividades y responsabilidades (ítems 13-16)
    {"id":13, "cat":2, "dom":1, "sec":"Actividades y responsabilidades",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"En mi trabajo soy responsable de cosas de mucho valor."},
    {"id":14, "cat":2, "dom":1, "sec":"Actividades y responsabilidades",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Respondo ante mi jefe por los resultados de toda mi área de trabajo."},
    {"id":15, "cat":2, "dom":1, "sec":"Actividades y responsabilidades",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"En el trabajo me dan órdenes contradictorias."},
    {"id":16, "cat":2, "dom":1, "sec":"Actividades y responsabilidades",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Considero que en mi trabajo me piden hacer cosas innecesarias."},

    # SECCIÓN 5: Jornada de trabajo (ítems 17-22)
    {"id":17, "cat":3, "dom":3, "sec":"Jornada de trabajo",
     "cat_nombre":"Organización del tiempo de trabajo",
     "dom_nombre":"Jornada de trabajo",
     "texto":"Trabajo horas extras más de tres veces a la semana."},
    {"id":18, "cat":3, "dom":3, "sec":"Jornada de trabajo",
     "cat_nombre":"Organización del tiempo de trabajo",
     "dom_nombre":"Jornada de trabajo",
     "texto":"Mi trabajo me exige laborar en días de descanso, festivos o fines de semana."},
    {"id":19, "cat":3, "dom":4, "sec":"Jornada de trabajo",
     "cat_nombre":"Organización del tiempo de trabajo",
     "dom_nombre":"Interferencia en la relación trabajo-familia",
     "texto":"Considero que el tiempo en el trabajo es mucho y perjudica mis actividades familiares o personales."},
    {"id":20, "cat":3, "dom":4, "sec":"Jornada de trabajo",
     "cat_nombre":"Organización del tiempo de trabajo",
     "dom_nombre":"Interferencia en la relación trabajo-familia",
     "texto":"Debo atender asuntos de trabajo cuando estoy en casa."},
    {"id":21, "cat":3, "dom":4, "sec":"Jornada de trabajo",
     "cat_nombre":"Organización del tiempo de trabajo",
     "dom_nombre":"Interferencia en la relación trabajo-familia",
     "texto":"Pienso en las actividades familiares o personales cuando estoy en mi trabajo."},
    {"id":22, "cat":3, "dom":4, "sec":"Jornada de trabajo",
     "cat_nombre":"Organización del tiempo de trabajo",
     "dom_nombre":"Interferencia en la relación trabajo-familia",
     "texto":"Pienso que mis responsabilidades familiares afectan mi trabajo."},

    # SECCIÓN 6: Decisiones en el trabajo (ítems 23-28) — INVERSOS
    {"id":23, "cat":2, "dom":2, "sec":"Decisiones en el trabajo",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Falta de control sobre el trabajo",
     "texto":"Mi trabajo permite que desarrolle nuevas habilidades."},
    {"id":24, "cat":2, "dom":2, "sec":"Decisiones en el trabajo",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Falta de control sobre el trabajo",
     "texto":"En mi trabajo puedo aspirar a un mejor puesto."},
    {"id":25, "cat":2, "dom":2, "sec":"Decisiones en el trabajo",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Falta de control sobre el trabajo",
     "texto":"Durante mi jornada de trabajo puedo tomar pausas cuando las necesito."},
    {"id":26, "cat":2, "dom":2, "sec":"Decisiones en el trabajo",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Falta de control sobre el trabajo",
     "texto":"Puedo decidir cuánto trabajo realizo durante la jornada laboral."},
    {"id":27, "cat":2, "dom":2, "sec":"Decisiones en el trabajo",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Falta de control sobre el trabajo",
     "texto":"Puedo decidir la velocidad a la que realizo mis actividades en mi trabajo."},
    {"id":28, "cat":2, "dom":2, "sec":"Decisiones en el trabajo",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Falta de control sobre el trabajo",
     "texto":"Puedo cambiar el orden de las actividades que realizo en mi trabajo."},

    # SECCIÓN 7: Cambios en el trabajo (ítems 29-30)
    {"id":29, "cat":2, "dom":2, "sec":"Cambios en el trabajo",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Falta de control sobre el trabajo",
     "texto":"Los cambios que se presentan en mi trabajo dificultan mi labor."},
    {"id":30, "cat":2, "dom":2, "sec":"Cambios en el trabajo",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Falta de control sobre el trabajo",
     "texto":"Cuando se presentan cambios en mi trabajo se tienen en cuenta mis ideas o aportaciones."},

    # SECCIÓN 8: Capacitación e información (ítems 31-36) — INVERSOS
    {"id":31, "cat":4, "dom":5, "sec":"Capacitación e información",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"Me informan con claridad cuáles son mis funciones."},
    {"id":32, "cat":4, "dom":5, "sec":"Capacitación e información",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"Me explican claramente los resultados que debo obtener en mi trabajo."},
    {"id":33, "cat":4, "dom":5, "sec":"Capacitación e información",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"Me explican claramente los objetivos de mi trabajo."},
    {"id":34, "cat":4, "dom":5, "sec":"Capacitación e información",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"Me informan con quién puedo resolver problemas o asuntos de trabajo."},
    {"id":35, "cat":2, "dom":2, "sec":"Capacitación e información",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Falta de control sobre el trabajo",
     "texto":"Me permiten asistir a capacitaciones relacionadas con mi trabajo."},
    {"id":36, "cat":2, "dom":2, "sec":"Capacitación e información",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Falta de control sobre el trabajo",
     "texto":"Recibo capacitación útil para hacer mi trabajo."},

    # SECCIÓN 9: El jefe (ítems 37-41) — INVERSOS
    {"id":37, "cat":4, "dom":5, "sec":"El jefe",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"Mi jefe ayuda a organizar mejor el trabajo."},
    {"id":38, "cat":4, "dom":5, "sec":"El jefe",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"Mi jefe tiene en cuenta mis puntos de vista y opiniones."},
    {"id":39, "cat":4, "dom":5, "sec":"El jefe",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"Mi jefe me comunica a tiempo la información relacionada con el trabajo."},
    {"id":40, "cat":4, "dom":5, "sec":"El jefe",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"La orientación que me da mi jefe me ayuda a realizar mejor mi trabajo."},
    {"id":41, "cat":4, "dom":5, "sec":"El jefe",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"Mi jefe ayuda a solucionar los problemas que se presentan en el trabajo."},

    # SECCIÓN 10: Relaciones con compañeros (ítems 42-46) — INVERSOS
    {"id":42, "cat":4, "dom":6, "sec":"Relaciones con compañeros",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Relaciones en el trabajo",
     "texto":"Puedo confiar en mis compañeros de trabajo."},
    {"id":43, "cat":4, "dom":6, "sec":"Relaciones con compañeros",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Relaciones en el trabajo",
     "texto":"Entre compañeros solucionamos los problemas de trabajo de forma respetuosa."},
    {"id":44, "cat":4, "dom":6, "sec":"Relaciones con compañeros",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Relaciones en el trabajo",
     "texto":"En mi trabajo me hacen sentir parte del grupo."},
    {"id":45, "cat":4, "dom":6, "sec":"Relaciones con compañeros",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Relaciones en el trabajo",
     "texto":"Cuando tenemos que realizar trabajo de equipo los compañeros colaboran."},
    {"id":46, "cat":4, "dom":6, "sec":"Relaciones con compañeros",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Relaciones en el trabajo",
     "texto":"Mis compañeros de trabajo me ayudan cuando tengo dificultades."},

    # SECCIÓN 11: Rendimiento, reconocimiento, pertenencia y estabilidad (ítems 47-56)
    {"id":47, "cat":5, "dom":8, "sec":"Rendimiento y reconocimiento",
     "cat_nombre":"Entorno organizacional",
     "dom_nombre":"Reconocimiento del desempeño",
     "texto":"Me informan sobre lo que hago bien en mi trabajo."},
    {"id":48, "cat":5, "dom":8, "sec":"Rendimiento y reconocimiento",
     "cat_nombre":"Entorno organizacional",
     "dom_nombre":"Reconocimiento del desempeño",
     "texto":"La forma como evalúan mi trabajo en mi centro de trabajo me ayuda a mejorar mi desempeño."},
    {"id":49, "cat":5, "dom":8, "sec":"Rendimiento y reconocimiento",
     "cat_nombre":"Entorno organizacional",
     "dom_nombre":"Reconocimiento del desempeño",
     "texto":"En mi centro de trabajo me pagan a tiempo mi salario."},
    {"id":50, "cat":5, "dom":8, "sec":"Rendimiento y reconocimiento",
     "cat_nombre":"Entorno organizacional",
     "dom_nombre":"Reconocimiento del desempeño",
     "texto":"El pago que recibo es el que merezco por el trabajo que realizo."},
    {"id":51, "cat":5, "dom":8, "sec":"Rendimiento y reconocimiento",
     "cat_nombre":"Entorno organizacional",
     "dom_nombre":"Reconocimiento del desempeño",
     "texto":"Si obtengo los resultados esperados en mi trabajo me recompensan o reconocen."},
    {"id":52, "cat":5, "dom":8, "sec":"Rendimiento y reconocimiento",
     "cat_nombre":"Entorno organizacional",
     "dom_nombre":"Reconocimiento del desempeño",
     "texto":"Las personas que hacen bien el trabajo pueden crecer laboralmente."},
    {"id":53, "cat":5, "dom":9, "sec":"Rendimiento y reconocimiento",
     "cat_nombre":"Entorno organizacional",
     "dom_nombre":"Insuficiente sentido de pertenencia e inestabilidad",
     "texto":"Considero que mi trabajo es estable."},
    {"id":54, "cat":5, "dom":9, "sec":"Rendimiento y reconocimiento",
     "cat_nombre":"Entorno organizacional",
     "dom_nombre":"Insuficiente sentido de pertenencia e inestabilidad",
     "texto":"En mi trabajo existe continua rotación de personal."},
    {"id":55, "cat":5, "dom":9, "sec":"Rendimiento y reconocimiento",
     "cat_nombre":"Entorno organizacional",
     "dom_nombre":"Insuficiente sentido de pertenencia e inestabilidad",
     "texto":"Siento orgullo de laborar en este centro de trabajo."},
    {"id":56, "cat":5, "dom":9, "sec":"Rendimiento y reconocimiento",
     "cat_nombre":"Entorno organizacional",
     "dom_nombre":"Insuficiente sentido de pertenencia e inestabilidad",
     "texto":"Me siento comprometido con mi trabajo."},

    # SECCIÓN 12: Violencia laboral (ítems 57-64)
    {"id":57, "cat":4, "dom":7, "sec":"Violencia laboral",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Violencia laboral",
     "texto":"En mi trabajo puedo expresarme libremente sin interrupciones."},
    {"id":58, "cat":4, "dom":7, "sec":"Violencia laboral",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Violencia laboral",
     "texto":"Recibo críticas constantes a mi persona y/o trabajo."},
    {"id":59, "cat":4, "dom":7, "sec":"Violencia laboral",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Violencia laboral",
     "texto":"Recibo burlas, calumnias, difamaciones, humillaciones o ridiculizaciones."},
    {"id":60, "cat":4, "dom":7, "sec":"Violencia laboral",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Violencia laboral",
     "texto":"Se ignora mi presencia o se me excluye de las reuniones de trabajo y en la toma de decisiones."},
    {"id":61, "cat":4, "dom":7, "sec":"Violencia laboral",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Violencia laboral",
     "texto":"Se manipulan las situaciones de trabajo para hacerme parecer un mal trabajador."},
    {"id":62, "cat":4, "dom":7, "sec":"Violencia laboral",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Violencia laboral",
     "texto":"Se ignoran mis éxitos laborales y se atribuyen a otros trabajadores."},
    {"id":63, "cat":4, "dom":7, "sec":"Violencia laboral",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Violencia laboral",
     "texto":"Me bloquean o impiden las oportunidades que tengo para obtener ascenso o mejora en mi trabajo."},
    {"id":64, "cat":4, "dom":7, "sec":"Violencia laboral",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Violencia laboral",
     "texto":"He presenciado actos de violencia en mi centro de trabajo."},

    # SECCIÓN 13: Atención a clientes (65-68) — CONDICIONAL
    {"id":65, "cat":2, "dom":1, "sec":"Atención a clientes",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Atiendo clientes o usuarios muy enojados."},
    {"id":66, "cat":2, "dom":1, "sec":"Atención a clientes",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Mi trabajo me exige atender personas muy necesitadas de ayuda o enfermas."},
    {"id":67, "cat":2, "dom":1, "sec":"Atención a clientes",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Para hacer mi trabajo debo demostrar sentimientos distintos a los míos."},
    {"id":68, "cat":2, "dom":1, "sec":"Atención a clientes",
     "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Mi trabajo me exige atender situaciones de violencia."},

    # SECCIÓN 14: Supervisión de trabajadores (69-72) — CONDICIONAL
    {"id":69, "cat":4, "dom":6, "sec":"Supervisión de trabajadores",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Relaciones en el trabajo",
     "texto":"Comunican tarde los asuntos de trabajo."},
    {"id":70, "cat":4, "dom":6, "sec":"Supervisión de trabajadores",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Relaciones en el trabajo",
     "texto":"Dificultan el logro de los resultados del trabajo."},
    {"id":71, "cat":4, "dom":6, "sec":"Supervisión de trabajadores",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Relaciones en el trabajo",
     "texto":"Cooperan poco cuando se necesita."},
    {"id":72, "cat":4, "dom":6, "sec":"Supervisión de trabajadores",
     "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Relaciones en el trabajo",
     "texto":"Ignoran las sugerencias para mejorar su trabajo."},
]

# ── Instrucciones por sección ─────────────────────────────────────────────────
INSTRUCCIONES_G3 = {
    "Condiciones ambientales":
        "Para responder las preguntas siguientes considere las condiciones ambientales de su centro de trabajo.",
    "Cantidad y ritmo de trabajo":
        "Para responder a las preguntas siguientes piense en la cantidad y ritmo de trabajo que tiene.",
    "Esfuerzo mental":
        "Las preguntas siguientes están relacionadas con el esfuerzo mental que le exige su trabajo.",
    "Actividades y responsabilidades":
        "Las preguntas siguientes están relacionadas con las actividades que realiza en su trabajo y las responsabilidades que tiene.",
    "Jornada de trabajo":
        "Las preguntas siguientes están relacionadas con su jornada de trabajo.",
    "Decisiones en el trabajo":
        "Las preguntas siguientes están relacionadas con las decisiones que puede tomar en su trabajo.",
    "Cambios en el trabajo":
        "Las preguntas siguientes están relacionadas con cualquier tipo de cambio que ocurra en su trabajo (considere los últimos cambios realizados).",
    "Capacitación e información":
        "Las preguntas siguientes están relacionadas con la capacitación e información que se le proporciona sobre su trabajo.",
    "El jefe":
        "Las preguntas siguientes están relacionadas con el o los jefes con quien tiene contacto.",
    "Relaciones con compañeros":
        "Las preguntas siguientes se refieren a las relaciones con sus compañeros.",
    "Rendimiento y reconocimiento":
        "Las preguntas siguientes están relacionadas con la información que recibe sobre su rendimiento en el trabajo, el reconocimiento, el sentido de pertenencia y la estabilidad que le ofrece su trabajo.",
    "Violencia laboral":
        "Las preguntas siguientes están relacionadas con actos de violencia laboral (malos tratos, acoso, hostigamiento, acoso psicológico).",
    "Atención a clientes":
        "Las preguntas siguientes están relacionadas con la atención a clientes y usuarios.",
    "Supervisión de trabajadores":
        "Las preguntas siguientes están relacionadas con las actitudes de las personas que supervisa.",
}

# ── Calificación ──────────────────────────────────────────────────────────────
def nivel_riesgo_g3(puntaje: int) -> dict:
    for pmin, pmax, nivel, color, bg in TABLA5_G3:
        if pmin <= puntaje <= pmax:
            return {"nivel": nivel, "color": color, "bg": bg}
    return {"nivel": "MUY ALTO", "color": "#A20000", "bg": "#FDDEDE"}

def nivel_dominio_g3(dom_id: int, puntaje: int) -> str:
    c = TABLA_DOM_G3.get(dom_id, (4, 8, 12, 16))
    if puntaje <= c[0]: return "NULO"
    if puntaje <= c[1]: return "BAJO"
    if puntaje <= c[2]: return "MEDIO"
    if puntaje <= c[3]: return "ALTO"
    return "MUY ALTO"

def calcular_puntaje_g3(respuestas: list) -> dict:
    respuestas = list(respuestas) + ["Nunca"] * max(0, 72 - len(respuestas))
    puntaje_total = 0
    por_dominio   = {i: 0 for i in range(10)}
    por_categoria = {i: 0 for i in range(1, 6)}
    alerta_violencia = False
    for preg in PREGUNTAS_G3:
        idx  = preg["id"] - 1
        resp = respuestas[idx] if idx < len(respuestas) else "Nunca"
        esc  = ESCALA_I_G3 if preg["id"] in ITEMS_INVERSOS_G3 else ESCALA_D_G3
        pts  = esc.get(resp, 0)
        puntaje_total            += pts
        por_dominio[preg["dom"]] += pts
        por_categoria[preg["cat"]] += pts
        if preg["dom"] == 7 and pts > 0:
            alerta_violencia = True
    nr = nivel_riesgo_g3(puntaje_total)
    return {
        "puntaje_total":    puntaje_total,
        "nivel":            nr["nivel"],
        "color":            nr["color"],
        "bg":               nr["bg"],
        "por_dominio":      por_dominio,
        "por_categoria":    por_categoria,
        "niveles_dom":      {d: nivel_dominio_g3(d, p) for d, p in por_dominio.items()},
        "alerta_violencia": alerta_violencia,
    }

# ── Excel ──────────────────────────────────────────────────────────────────────
def init_excel(path: str):
    os.makedirs("data", exist_ok=True)
    if not os.path.exists(path):
        cols = ["Folio","Fecha","Cliente","Razón Social","Nombre","Sexo","Edad",
                "Estado Civil","Nivel Estudios","Estud. Status","Puesto","Área",
                "Contratación","Tipo Personal","Jornada","Rotación Turnos",
                "Tiempo Puesto","Experiencia",
                "Puntaje Total","Nivel Riesgo",
                "Cat1 Ambiente","Cat2 Actividad","Cat3 Tiempo",
                "Cat4 Liderazgo","Cat5 Entorno",
                "Dom0 Ambiente","Dom1 Carga","Dom2 Control","Dom3 Jornada",
                "Dom4 Interferencia","Dom5 Liderazgo","Dom6 Relaciones",
                "Dom7 Violencia","Dom8 Reconocimiento","Dom9 Pertenencia",
                "Nivel Dom0","Nivel Dom1","Nivel Dom2","Nivel Dom3","Nivel Dom4",
                "Nivel Dom5","Nivel Dom6","Nivel Dom7","Nivel Dom8","Nivel Dom9",
                "Alerta Violencia"]
        for i in range(1, 73):
            cols.append(f"P{i:02d}")
        pd.DataFrame(columns=cols).to_csv(path, index=False, encoding="utf-8-sig")

def guardar(data: dict) -> bool:
    path = excel_path(data["cliente"], data.get("razon", ""))
    init_excel(path)
    res  = data.get("resultado", {})
    for intento in range(5):
        try:
            df   = pd.read_csv(path, encoding="utf-8-sig")
            fila = {
                "Folio":          data["folio"],
                "Fecha":          datetime.now().strftime("%Y-%m-%d %H:%M"),
                "Cliente":        data["cliente"],
                "Razón Social":   data["razon"],
                "Nombre":         data["nombre"],
                "Sexo":           data["sexo"],
                "Edad":           data["edad"],
                "Estado Civil":   data["ecivil"],
                "Nivel Estudios": data["estudios"],
                "Estud. Status":  data["estatus"],
                "Puesto":         data["puesto"],
                "Área":           data["area"],
                "Contratación":   data["contrat"],
                "Tipo Personal":  data["personal"],
                "Jornada":        data["jornada"],
                "Rotación Turnos":data["rotacion"],
                "Tiempo Puesto":  data["tpuesto"],
                "Experiencia":    data["exp"],
                "Puntaje Total":  res.get("puntaje_total", 0),
                "Nivel Riesgo":   res.get("nivel", ""),
                "Cat1 Ambiente":  res.get("por_categoria", {}).get(1, 0),
                "Cat2 Actividad": res.get("por_categoria", {}).get(2, 0),
                "Cat3 Tiempo":    res.get("por_categoria", {}).get(3, 0),
                "Cat4 Liderazgo": res.get("por_categoria", {}).get(4, 0),
                "Cat5 Entorno":   res.get("por_categoria", {}).get(5, 0),
                "Dom0 Ambiente":  res.get("por_dominio", {}).get(0, 0),
                "Dom1 Carga":     res.get("por_dominio", {}).get(1, 0),
                "Dom2 Control":   res.get("por_dominio", {}).get(2, 0),
                "Dom3 Jornada":   res.get("por_dominio", {}).get(3, 0),
                "Dom4 Interferencia": res.get("por_dominio", {}).get(4, 0),
                "Dom5 Liderazgo": res.get("por_dominio", {}).get(5, 0),
                "Dom6 Relaciones":res.get("por_dominio", {}).get(6, 0),
                "Dom7 Violencia": res.get("por_dominio", {}).get(7, 0),
                "Dom8 Reconocimiento": res.get("por_dominio", {}).get(8, 0),
                "Dom9 Pertenencia":    res.get("por_dominio", {}).get(9, 0),
                "Nivel Dom0": res.get("niveles_dom", {}).get(0, ""),
                "Nivel Dom1": res.get("niveles_dom", {}).get(1, ""),
                "Nivel Dom2": res.get("niveles_dom", {}).get(2, ""),
                "Nivel Dom3": res.get("niveles_dom", {}).get(3, ""),
                "Nivel Dom4": res.get("niveles_dom", {}).get(4, ""),
                "Nivel Dom5": res.get("niveles_dom", {}).get(5, ""),
                "Nivel Dom6": res.get("niveles_dom", {}).get(6, ""),
                "Nivel Dom7": res.get("niveles_dom", {}).get(7, ""),
                "Nivel Dom8": res.get("niveles_dom", {}).get(8, ""),
                "Nivel Dom9": res.get("niveles_dom", {}).get(9, ""),
                "Alerta Violencia": "SÍ — URGENTE" if res.get("alerta_violencia") else "No",
            }
            for i, resp in enumerate(data.get("respuestas", []), 1):
                fila[f"P{i:02d}"] = resp
            df = pd.concat([df, pd.DataFrame([fila])], ignore_index=True)
            df.to_csv(path, index=False, encoding="utf-8-sig")
            return True
        except PermissionError:
            if intento < 4: time.sleep(1)
            else:
                st.error("⚠ El archivo Excel está abierto. Ciérralo e intenta de nuevo.")
                return False
    return False

# ── Folio consecutivo basado en CSV ──────────────────────────────────────────
_folio_lock = threading.Lock()

def folio_nuevo(cliente_key: str, razon_social: str) -> str:
    """Folio 1, 2, 3... basado en registros del CSV."""
    path      = excel_path(cliente_key, razon_social)
    razon_key = razon_social.strip().upper()
    with _folio_lock:
        init_excel(path)
        try:
            df = pd.read_csv(path, encoding="utf-8-sig")
            if df.empty or "Razón Social" not in df.columns:
                n = 1
            else:
                mask = df["Razón Social"].astype(str).str.strip().str.upper() == razon_key
                n    = int(mask.sum()) + 1
        except Exception:
            n = 1
    return str(n)

def idx_de(lst, val): return lst.index(val) if val in lst else 0
def solo_letras(t): return re.sub(r"[^A-Za-záéíóúÁÉÍÓÚüÜñÑ\s;]", "", t).upper().strip()

def trabajador_ya_registrado(ap1, ap2, nom, razon, cliente_key) -> bool:
    path = excel_path(cliente_key, razon)
    if not os.path.exists(path): return False
    try:
        df = pd.read_csv(path, encoding="utf-8-sig")
        if df.empty or "Nombre" not in df.columns: return False
        nombre_nuevo = f"{ap1}; {ap2}; {nom}".strip().upper()
        mask = ((df["Nombre"].str.strip().str.upper() == nombre_nuevo) &
                (df["Razón Social"].str.strip().str.upper() == razon.strip().upper()))
        return mask.any()
    except: return False

def _img_b64(path: str) -> str:
    import base64, mimetypes
    try:
        mime = mimetypes.guess_type(path)[0] or "image/png"
        with open(path, "rb") as f:
            return f"data:{mime};base64,{base64.b64encode(f.read()).decode()}"
    except: return ""

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700&display=swap');
:root{--v:#4b694e;--v2:#6a9370;--rj:#a20000;--fo:#f1f1f1;--bl:#ffffff;
      --tx:#1a1a1a;--br:#d0d0d0;--am:#c8a600;}
html,body,.stApp{background:var(--fo)!important;font-family:'Montserrat',sans-serif!important;}
#MainMenu,footer,header{visibility:hidden;}
.block-container{padding:1.6rem 2rem!important;max-width:860px!important;margin:auto;}
.ptitulo{font-size:1.35rem;font-weight:700;color:var(--v);border-bottom:2px solid var(--v);
         padding-bottom:.4rem;margin-bottom:1.2rem;}
.pcard{background:var(--bl);border-radius:13px;padding:1.4rem 1.8rem;
       margin-bottom:1rem;box-shadow:0 2px 8px rgba(0,0,0,.07);}
.bv-box{background:var(--bl);border-radius:15px;padding:2rem 2.5rem;
        text-align:center;box-shadow:0 4px 14px rgba(0,0,0,.09);margin-bottom:1.4rem;}
.bv-tit{font-size:1.55rem;font-weight:700;color:var(--v);text-transform:uppercase;
        letter-spacing:.05em;margin-bottom:.3rem;}
.bv-sub{font-size:.88rem;font-weight:500;color:#666;text-transform:uppercase;
        letter-spacing:.1em;margin-bottom:1.2rem;}
.priv{background:#f8f8f8;border:1px solid var(--br);border-radius:9px;
      padding:1rem 1.4rem;text-align:left;font-size:.92rem;color:#555;
      margin-bottom:1.2rem;line-height:1.8;}
.slabel{font-size:.78rem;font-weight:700;color:var(--v);text-transform:uppercase;
        letter-spacing:.1em;margin:.9rem 0 .25rem 0;}
.cfm-box{background:var(--bl);border:2px solid var(--v);border-radius:13px;
         padding:2rem 2.4rem;text-align:center;font-size:1rem;line-height:1.9;
         color:var(--tx);margin-bottom:1.4rem;box-shadow:0 4px 14px rgba(75,105,78,.14);}
.av-box{background:var(--bl);border:2px solid var(--am);border-radius:13px;
        padding:1.7rem 2.2rem;text-align:center;font-size:1rem;color:var(--tx);
        line-height:1.9;margin-bottom:1.4rem;}
.av-tit{font-size:1.05rem;font-weight:700;color:#7a5900;margin-bottom:.5rem;}
.pq-card{background:var(--bl);border-radius:13px;padding:1.7rem 2.1rem;
         margin-bottom:.85rem;box-shadow:0 3px 11px rgba(0,0,0,.08);}
.pq-dom{font-size:.72rem;color:#888;margin-bottom:.2rem;}
.pq-txt{font-size:1.05rem;font-weight:600;color:var(--tx);line-height:1.6;}
.pq-num{font-size:.7rem;color:#aaa;margin-bottom:.3rem;}
.err-r{color:var(--rj);font-size:.88rem;font-weight:700;margin-top:.4rem;}
.prog-txt{font-size:.82rem;color:#888;text-align:right;margin-bottom:.5rem;font-weight:500;}
.fin-box{background:var(--v);color:#fff;border-radius:17px;padding:3rem 2rem;
         text-align:center;box-shadow:0 6px 22px rgba(75,105,78,.24);}
.fin-tit{font-size:1.35rem;font-weight:700;letter-spacing:.05em;margin-bottom:.6rem;}
.fin-sub{font-size:.96rem;opacity:.84;}
.dup-alert{background:#fff0f0;border:2px solid #a20000;border-radius:11px;
           padding:1rem 1.4rem;margin:.8rem 0;font-size:.88rem;color:#a20000;
           font-weight:600;line-height:1.7;}
.stButton>button{font-family:'Montserrat',sans-serif!important;font-weight:700!important;
  font-size:.93rem!important;letter-spacing:.05em!important;border-radius:9px!important;
  border:none!important;padding:.65rem 1.8rem!important;cursor:pointer!important;
  background:var(--v)!important;color:#fff!important;transition:opacity .2s!important;}
.stButton>button:hover{opacity:.84!important;}
div[role="radiogroup"]>label{background:#fafafa;border:1.5px solid var(--br);
  border-radius:7px;padding:.5rem 1rem!important;margin-bottom:.35rem!important;
  font-size:.96rem!important;font-weight:500;cursor:pointer;
  transition:border-color .14s,background .14s;}
div[role="radiogroup"]>label:hover{border-color:var(--v2);background:#f0f6f0;}
hr.div{border:none;border-top:1.5px solid var(--br);margin:.8rem 0;}
.stTextInput input{text-transform:uppercase!important;font-family:Montserrat,sans-serif!important;
  font-weight:600!important;letter-spacing:.04em!important;}
@keyframes rf-pulse{0%,100%{opacity:1;transform:scale(1);}50%{opacity:.55;transform:scale(.96);}}
@keyframes rf-spin{0%{transform:rotate(0deg);}100%{transform:rotate(360deg);}}
.loading-overlay{position:fixed;inset:0;background:rgba(241,241,241,.92);
  display:flex;flex-direction:column;align-items:center;justify-content:center;z-index:9999;}
.loading-logo{animation:rf-pulse 1.4s ease-in-out infinite;}
.loading-ring{width:48px;height:48px;border:4px solid #d0d0d0;border-top-color:var(--v);
  border-radius:50%;animation:rf-spin .9s linear infinite;margin-top:1.2rem;}
.loading-txt{font-size:.82rem;font-weight:600;color:#4b694e;margin-top:.7rem;letter-spacing:.06em;}
</style>
""", unsafe_allow_html=True)

# ── Estado de sesión ───────────────────────────────────────────────────────────
_cliente_def  = _CLIENTE_URL if (_MODO_EMPLEADO and _CLIENTE_URL in CLIENTES) else "FRUCO"
_pantalla_def = "bienvenida" if _MODO_EMPLEADO else "panel"

# ── Modo empleado: forzar bienvenida al entrar con ?cliente= ─────────────────
if _MODO_EMPLEADO:
    _session_cliente = st.session_state.get("_session_cliente_key", "")
    _pantalla_actual = st.session_state.get("pantalla", "")
    if _session_cliente != _CLIENTE_URL:
        st.session_state["_session_cliente_key"] = _CLIENTE_URL
        if _pantalla_actual in ["panel", "", None]:
            st.session_state["pantalla"]    = "bienvenida"
            st.session_state["cliente_key"] = _cliente_def
            st.session_state["razon"]       = CLIENTES[_cliente_def]["opciones"][0]

DEF = dict(
    pantalla=_pantalla_def, cliente_key=_cliente_def,
    razon=CLIENTES[_cliente_def]["opciones"][0],
    areas=["Producción","Administración","Recursos Humanos","Ventas","Operaciones"],
    folio="001",
    ap1="", ap2="", nom="",
    sexo=SEL, edad=SEL, ecivil=SEL,
    estudios=SEL, estatus="Terminada",
    puesto=SEL, area=SEL,
    contrat=SEL, personal=SEL,
    jornada=SEL, rotacion="Sí",
    tpuesto=SEL, exp=SEL,
    preg_idx=0, respuestas=[],
    atiende_clientes=None,
    es_jefe=None,
    err=False, res=None, modal=None,
    form_v=0,
)
for k, v in DEF.items():
    if k not in st.session_state:
        st.session_state[k] = v
S = st.session_state

# Seguridad: si pantalla=panel en modo empleado, corregir
if _MODO_EMPLEADO and S.get("pantalla") == "panel":
    S["pantalla"] = "bienvenida"
KEYS_FORM = ["d_ap1","d_ap2","d_nom","w_sexo","w_edad","w_ecivil","w_estudios",
             "w_estatus","w_puesto","w_area","w_contrat","w_personal",
             "w_jornada","w_rotacion","w_tpuesto","w_exp"]

def limpiar():
    for k in ["ap1","ap2","nom"]: S[k] = ""
    for k in ["sexo","edad","ecivil","estudios","puesto","area",
              "contrat","personal","jornada","tpuesto","exp"]: S[k] = SEL
    S.estatus = "Terminada"; S.rotacion = "Sí"
    for wk in KEYS_FORM:
        if wk in st.session_state: del st.session_state[wk]

def borrar_formulario():
    limpiar(); S.form_v = S.get("form_v", 0) + 1

def reset_cuestionario():
    S.preg_idx = 0; S.respuestas = []
    S.atiende_clientes = None; S.es_jefe = None
    S.err = False; S.res = None

def header(mostrar_folio=False):
    rf_src  = _img_b64(LOGO_RF)
    cli_src = _img_b64(CLIENTES[S.cliente_key]["logo"])
    folio_html = (
        f'<div style="font-size:.78rem;font-weight:600;color:#888;text-align:right;'
        f'letter-spacing:.06em;text-transform:uppercase;line-height:1.4;">'
        f'NO. DE CUESTIONARIO<br>'
        f'<span style="font-size:1.25rem;font-weight:700;color:#1a1a1a;">{S.folio}</span></div>'
    ) if (mostrar_folio and not _MODO_EMPLEADO) else ""
    rf_tag  = (f'<img src="{rf_src}" alt="RFRANYUTTI" '
               f'style="height:52px;max-width:120px;width:auto;object-fit:contain;">'
               ) if rf_src else "<b>RFRANYUTTI</b>"
    cli_tag = (f'<img src="{cli_src}" alt="{S.cliente_key}" '
               f'style="height:52px;max-width:120px;width:auto;object-fit:contain;">'
               ) if cli_src else f"<b>{S.cliente_key}</b>"
    st.markdown(f"""
    <div style="display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;
                gap:.5rem;padding:.5rem 0 .6rem 0;margin-bottom:.2rem;">
        <div style="display:flex;align-items:center;gap:.8rem;flex-wrap:wrap;">
            {rf_tag}
            <div style="width:1px;height:44px;background:#ccc;flex-shrink:0;"></div>
            {cli_tag}
        </div>
        <div style="flex-shrink:0;">{folio_html}</div>
    </div>
    <hr class="div">
    """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PANEL DEL OPERATIVO
# ══════════════════════════════════════════════════════════════════════════════
if S.pantalla == "panel":
    st.markdown('<div class="ptitulo">⚙ PANEL DEL OPERATIVO · NOM-035-STPS-2018 · GUÍA III</div>',
                unsafe_allow_html=True)

    if S.modal == "borrar":
        st.warning("⚠ **¿Eliminar todos los registros?** Esta acción no se puede deshacer.")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("CONTINUAR — BORRAR TODO"):
                p = excel_path(S.cliente_key, S.razon)
                if os.path.exists(p): os.remove(p)
                S.modal = None; st.success("✓ Registros eliminados."); st.rerun()
        with c2:
            if st.button("CANCELAR"): S.modal = None; st.rerun()
        st.stop()

    if S.modal == "terminar":
        c1, c2 = st.columns(2)
        with c1:
            if st.button("CONFIRMAR"):
                S.modal = None; st.success("✓ Sesión cerrada."); st.rerun()
        with c2:
            if st.button("CANCELAR"): S.modal = None; st.rerun()
        st.stop()

    st.markdown('<div class="pcard">', unsafe_allow_html=True)
    st.markdown("**1. Selección de cliente**")
    ck = st.selectbox("Cliente", list(CLIENTES.keys()),
                      index=list(CLIENTES.keys()).index(S.cliente_key))
    S.cliente_key = ck
    inf = CLIENTES[ck]
    S.razon = st.selectbox("Razón social", inf["opciones"])
    try: st.image(inf["logo"], width=120)
    except: st.caption(f"Logo: {inf['logo']}")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="pcard">', unsafe_allow_html=True)
    st.markdown("**2. Departamentos / Áreas**")
    at = st.text_area("", "\n".join(S.areas), height=120, label_visibility="collapsed")
    S.areas = [a.strip() for a in at.split("\n") if a.strip()]
    st.caption(f"{len(S.areas)} área(s) configuradas.")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="pcard">', unsafe_allow_html=True)
    st.markdown("**3. Guía a aplicar**")
    st.info("**GUÍA III** — Identificación y Análisis de los Factores de Riesgo Psicosocial "
            "y Evaluación del Entorno Organizacional\n\n"
            "Para centros de trabajo con **más de 50 trabajadores** · **72 preguntas**")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("**4. Acciones**")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        _p_raw = excel_path(S.cliente_key, S.razon)
        if os.path.exists(_p_raw):
            with open(_p_raw, "rb") as _f:
                st.download_button(
                    label="⬇️ DATOS (Excel)",
                    data=_f.read(),
                    file_name=os.path.basename(_p_raw),
                    mime="text/csv",
                    use_container_width=True)
        else:
            st.button("⬇️ DATOS (Excel)", disabled=True, use_container_width=True)
    with c2:
        st.button("ABRIR WORD", use_container_width=True)
    with c3:
        if st.button("BORRAR REGISTROS", use_container_width=True):
            S.modal = "borrar"; st.rerun()
    with c4:
        if st.button("TERMINAR REGISTROS", use_container_width=True):
            S.modal = "terminar"; st.rerun()

    st.markdown('<hr class="div">', unsafe_allow_html=True)
    st.markdown("**5. Generar Reportes**")
    _p_rep   = excel_path(S.cliente_key, S.razon)
    _hay_dat = os.path.exists(_p_rep)
    cr1, cr2 = st.columns(2)
    with cr1:
        if st.button("📊  EXCEL MEJORADO + GRÁFICAS", use_container_width=True,
                     disabled=not _hay_dat):
            with st.spinner("Generando Excel..."):
                try:
                    from generar_reporte_g3 import generar_excel_g3
                    out = generar_excel_g3(_p_rep, S.cliente_key, S.razon)
                    if out:
                        with open(out, "rb") as f:
                            st.download_button("⬇️  DESCARGAR EXCEL", data=f.read(),
                                file_name=os.path.basename(out),
                                mime="text/csv",
                                use_container_width=True)
                    else: st.warning("Sin datos suficientes.")
                except Exception as e: st.error(f"Error: {e}")
    with cr2:
        if st.button("📄  INFORME WORD COMPLETO", use_container_width=True,
                     disabled=not _hay_dat):
            with st.spinner("Generando Word..."):
                try:
                    from generar_reporte_g3 import generar_word_g3
                    out = generar_word_g3(_p_rep, S.cliente_key, S.razon,
                                          logo_rf=LOGO_RF,
                                          logo_cliente=CLIENTES[S.cliente_key]["logo"])
                    if out:
                        with open(out, "rb") as f:
                            st.download_button("⬇️  DESCARGAR WORD", data=f.read(),
                                file_name=os.path.basename(out),
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True)
                    else: st.warning("Sin datos suficientes.")
                except Exception as e: st.error(f"Error: {e}")
    if not _hay_dat:
        st.caption("⚠ Registra al menos un cuestionario antes de generar reportes.")

    st.markdown('<hr class="div">', unsafe_allow_html=True)
    st.markdown("**6. Link para empleados**")
    _base = st.text_input("URL base de tu app", value="http://localhost:8501",
                          help="Pega tu URL de Streamlit Cloud aquí.")
    _link = f"{_base.rstrip('/')}/?cliente={S.cliente_key}"
    st.markdown(f"""
    <div style="background:#f0f6f0;border:2px solid #4b694e;border-radius:10px;
                padding:1rem 1.4rem;margin:.4rem 0 .8rem 0;font-family:Montserrat,sans-serif;">
        <div style="font-size:.72rem;font-weight:700;color:#4b694e;text-transform:uppercase;
                    letter-spacing:.08em;margin-bottom:.35rem;">
            🔗 Link empleados — {S.cliente_key}
        </div>
        <div style="font-size:.9rem;font-weight:600;word-break:break-all;background:#fff;
                    border-radius:6px;padding:.45rem .8rem;border:1px solid #d0d0d0;">
            {_link}
        </div>
        <div style="font-size:.75rem;color:#666;margin-top:.4rem;">
            Comparte por WhatsApp · Sesiones independientes · Folio automático
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.code(_link, language=None)

    st.markdown('<hr class="div">', unsafe_allow_html=True)
    st.markdown("**7. Cuestionario modo local**")
    if st.button("🟢  INICIAR CUESTIONARIO PARA EMPLEADO", use_container_width=True):
        S.folio = folio_nuevo(S.cliente_key, S.razon)
        limpiar(); reset_cuestionario()
        S.pantalla = "bienvenida"; st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# BIENVENIDA
# ══════════════════════════════════════════════════════════════════════════════
elif S.pantalla == "bienvenida":
    header(False)
    st.markdown("""
    <div class="bv-box">
      <div class="bv-tit">NOM-035-STPS-2018</div>
      <div class="bv-sub">Cuestionario para identificar y analizar los factores de riesgo<br>
        psicosocial y evaluar el entorno organizacional</div>
      <p style="font-size:.83rem;color:#444;line-height:1.8;text-align:left;">
        Este cuestionario tiene como objetivo identificar y analizar los posibles factores
        de riesgo psicosocial en su centro de trabajo, así como evaluar el entorno
        organizacional. Sus respuestas son confidenciales y se utilizarán exclusivamente
        para mejorar las condiciones de trabajo. No existen respuestas correctas o
        incorrectas. Por favor responda con sinceridad considerando las condiciones
        de los <strong>dos últimos meses</strong>.
      </p>
      <div class="bv-sub" style="margin-top:.9rem;">Cláusula de Privacidad</div>
      <div class="priv">
        La información proporcionada será tratada con estricta confidencialidad. Los datos
        y resultados se utilizarán exclusivamente para fines internos de diagnóstico y mejora
        del ambiente laboral, conforme a la normativa vigente en materia de protección de
        datos personales.
      </div>
    </div>
    """, unsafe_allow_html=True)
    acepta = st.checkbox("He leído y acepto la cláusula de privacidad *")
    if acepta:
        if st.button("CONTINUAR →", use_container_width=True):
            if _MODO_EMPLEADO:
                S.folio = folio_nuevo(S.cliente_key, S.razon)
                limpiar(); reset_cuestionario()
            S.pantalla = "datos"; st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# DATOS GENERALES
# ══════════════════════════════════════════════════════════════════════════════
elif S.pantalla == "datos":
    header(True)
    st.markdown('<div class="slabel">Información General del Trabajador</div>',
                unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1: ap1 = st.text_input("PRIMER APELLIDO",  value=S.ap1, key=f"d_ap1_{S.form_v}")
    with c2: ap2 = st.text_input("SEGUNDO APELLIDO", value=S.ap2, key=f"d_ap2_{S.form_v}")
    with c3: nom = st.text_input("NOMBRE(S)",        value=S.nom, key=f"d_nom_{S.form_v}")

    _hay_err = False
    for _val, _lbl in [(ap1,"Primer Apellido"),(ap2,"Segundo Apellido"),(nom,"Nombre(s)")]:
        if _val and re.search(r"[^A-Za-záéíóúÁÉÍÓÚüÜñÑ\s]", _val):
            st.markdown(f'<p style="color:#a20000;font-size:.78rem;font-weight:700;">'
                        f'⚠ {_lbl}: solo letras.</p>', unsafe_allow_html=True)
            _hay_err = True

    c1,c2,c3 = st.columns(3)
    with c1: sexo   = st.selectbox("SEXO",      OPC_SEXO,   key=f"w_sexo_{S.form_v}",   index=idx_de(OPC_SEXO,S.sexo))
    with c2: edad   = st.selectbox("EDAD",      OPC_EDAD,   key=f"w_edad_{S.form_v}",   index=idx_de(OPC_EDAD,S.edad))
    with c3: ecivil = st.selectbox("ESTADO CIVIL", OPC_ECIVIL, key=f"w_ecivil_{S.form_v}", index=idx_de(OPC_ECIVIL,S.ecivil))

    c1,c2 = st.columns([2,1])
    with c1: estudios = st.selectbox("NIVEL DE ESTUDIOS", OPC_ESTUD, key=f"w_estudios_{S.form_v}", index=idx_de(OPC_ESTUD,S.estudios))
    with c2:
        if estudios not in [SEL,"Sin formación"]:
            estatus = st.radio("", ["Terminada","Incompleta"], horizontal=True,
                               key=f"w_estatus_{S.form_v}",
                               index=0 if S.estatus=="Terminada" else 1)
        else: estatus = "N/A"

    c1,c2 = st.columns(2)
    with c1: puesto = st.selectbox("PUESTO", OPC_PUESTO, key=f"w_puesto_{S.form_v}", index=idx_de(OPC_PUESTO,S.puesto))
    with c2:
        aopts = [SEL] + S.areas
        area  = st.selectbox("ÁREA", aopts, key=f"w_area_{S.form_v}", index=idx_de(aopts,S.area))

    c1,c2 = st.columns(2)
    with c1: contrat  = st.selectbox("TIPO CONTRATACIÓN", OPC_CONTRAT, key=f"w_contrat_{S.form_v}", index=idx_de(OPC_CONTRAT,S.contrat))
    with c2: personal = st.selectbox("TIPO PERSONAL",     OPC_PERSONAL, key=f"w_personal_{S.form_v}", index=idx_de(OPC_PERSONAL,S.personal))

    c1,c2 = st.columns([3,1])
    with c1: jornada  = st.selectbox("JORNADA", OPC_JORNADA, key=f"w_jornada_{S.form_v}", index=idx_de(OPC_JORNADA,S.jornada))
    with c2:
        st.markdown('<div class="slabel">Rotación</div>', unsafe_allow_html=True)
        rotacion = st.radio("", ["Sí","No"], horizontal=True, key=f"w_rotacion_{S.form_v}",
                            index=0 if S.rotacion=="Sí" else 1)

    c1,c2 = st.columns(2)
    with c1: tpuesto = st.selectbox("TIEMPO EN EL PUESTO", OPC_TPUESTO, key=f"w_tpuesto_{S.form_v}", index=idx_de(OPC_TPUESTO,S.tpuesto))
    with c2: exp     = st.selectbox("EXPERIENCIA LABORAL", OPC_EXP,     key=f"w_exp_{S.form_v}",     index=idx_de(OPC_EXP,S.exp))

    st.markdown("<br>", unsafe_allow_html=True)
    c1,c2 = st.columns([3,1])
    with c1: iniciar = st.button("INICIAR CUESTIONARIO", use_container_width=True)
    with c2:
        if st.button("BORRAR DATOS", use_container_width=True):
            borrar_formulario(); st.rerun()

    if iniciar:
        err = []
        a1=solo_letras(ap1.strip()); a2=solo_letras(ap2.strip()); nn=solo_letras(nom.strip())
        if _hay_err: err.append("Corrige los campos de nombre: solo letras.")
        else:
            if not a1: err.append("El Primer Apellido es obligatorio.")
            if not a2: err.append("El Segundo Apellido es obligatorio.")
            if not nn: err.append("El Nombre(s) es obligatorio.")
        for v,lbl in [(sexo,"Sexo"),(edad,"Edad"),(ecivil,"Estado Civil"),
                      (estudios,"Nivel de Estudios"),(puesto,"Puesto"),(area,"Área"),
                      (contrat,"Tipo de Contratación"),(personal,"Tipo de Personal"),
                      (jornada,"Jornada"),(tpuesto,"Tiempo en el Puesto"),(exp,"Experiencia")]:
            if v == SEL: err.append(f"Selecciona {lbl}.")
        if err:
            for e in err:
                st.markdown(f'<p style="color:#a20000;font-size:.82rem;">⚠ {e}</p>',
                            unsafe_allow_html=True)
        else:
            if trabajador_ya_registrado(a1, a2, nn, S.razon, S.cliente_key):
                st.markdown(f'<div class="dup-alert">⛔ {a1} {a2}, {nn} ya tiene '
                            f'cuestionario registrado.</div>', unsafe_allow_html=True)
            else:
                S.ap1=a1; S.ap2=a2; S.nom=nn
                S.sexo=sexo; S.edad=edad; S.ecivil=ecivil
                S.estudios=estudios; S.estatus=estatus
                S.puesto=puesto; S.area=area
                S.contrat=contrat; S.personal=personal
                S.jornada=jornada; S.rotacion=rotacion
                S.tpuesto=tpuesto; S.exp=exp
                S.pantalla = "confirmar"; st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# CONFIRMACIÓN
# ══════════════════════════════════════════════════════════════════════════════
elif S.pantalla == "confirmar":
    header(True)
    st.markdown("""
    <div class="cfm-box">
      <p style="font-size:.97rem;font-weight:700;text-transform:uppercase;color:#4b694e;">
        Confirmación de Datos</p>
      <p>Al pulsar <strong>"ACEPTAR"</strong> usted declara que los datos ingresados
      son correctos y reflejan su situación actual en la empresa.</p>
    </div>
    """, unsafe_allow_html=True)
    c1,c2 = st.columns(2)
    with c1:
        if st.button("ACEPTAR", use_container_width=True):
            S.pantalla = "aviso"; st.rerun()
    with c2:
        if st.button("← REGRESAR Y EDITAR", use_container_width=True):
            S.pantalla = "datos"; st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# AVISO
# ══════════════════════════════════════════════════════════════════════════════
elif S.pantalla == "aviso":
    header(True)
    st.markdown("""
    <div class="av-box">
      <div class="av-tit">⚠ ANTES DE COMENZAR — LEA CON ATENCIÓN</div>
      <p>Este cuestionario consta de <strong>72 preguntas</strong>. Una vez que avance
      a la siguiente pregunta, <strong>no podrá regresar a la anterior.</strong><br>
      Por favor responda con sinceridad considerando las condiciones de los
      <strong>dos últimos meses</strong>.</p>
    </div>
    """, unsafe_allow_html=True)
    if st.button("COMENZAR CUESTIONARIO →", use_container_width=True):
        reset_cuestionario()
        S.pantalla = "preguntas"; st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# PREGUNTAS
# ══════════════════════════════════════════════════════════════════════════════
elif S.pantalla == "preguntas":
    header(True)

    PREG_BASE = [p for p in PREGUNTAS_G3 if p["id"] <= 64]
    PREG_CLIE = [p for p in PREGUNTAS_G3 if p["id"] in [65,66,67,68]]
    PREG_JEFE = [p for p in PREGUNTAS_G3 if p["id"] in [69,70,71,72]]

    idx = S.preg_idx

    # ── Construir lista activa con lo que se sabe hasta ahora ─────────────────
    preg_activas = PREG_BASE[:]
    if S.atiende_clientes:  preg_activas += PREG_CLIE
    if S.es_jefe:           preg_activas += PREG_JEFE

    # ── Pregunta condicional: ¿atiende clientes? (después de pregunta 64) ─────
    if idx == 64 and S.atiende_clientes is None:
        st.progress(64/72)
        st.markdown('<div class="prog-txt">Pregunta 64 de 72 completada</div>',
                    unsafe_allow_html=True)
        st.markdown("""
        <div class="pq-card">
          <div class="pq-dom">Sección: Atención a clientes y usuarios</div>
          <div class="pq-txt">¿En mi trabajo debo brindar servicio a clientes o usuarios?</div>
        </div>
        """, unsafe_allow_html=True)
        c1,c2 = st.columns(2)
        with c1:
            if st.button("✅  SÍ, atiendo clientes", use_container_width=True):
                S.atiende_clientes = True; st.rerun()
        with c2:
            if st.button("❌  NO atiendo clientes", use_container_width=True):
                S.atiende_clientes = False; st.rerun()
        st.stop()

    # ── Pregunta condicional: ¿es jefe? ──────────────────────────────────────
    if S.atiende_clientes is not None and S.es_jefe is None:
        preg_hasta_jefe = PREG_BASE[:]
        if S.atiende_clientes: preg_hasta_jefe += PREG_CLIE
        if idx >= len(preg_hasta_jefe):
            st.progress(len(preg_hasta_jefe) / 72)
            st.markdown("""
            <div class="pq-card">
              <div class="pq-dom">Sección: Supervisión de trabajadores</div>
              <div class="pq-txt">¿Soy jefe de otros trabajadores?</div>
            </div>
            """, unsafe_allow_html=True)
            c1,c2 = st.columns(2)
            with c1:
                if st.button("✅  SÍ, soy jefe", use_container_width=True):
                    S.es_jefe = True
                    st.rerun()
            with c2:
                if st.button("❌  NO soy jefe", use_container_width=True):
                    S.es_jefe = False
                    resultado = calcular_puntaje_g3(S.respuestas + ["Nunca","Nunca","Nunca","Nunca"])
                    S.res = resultado
                    guardar(dict(
                        folio=S.folio, cliente=S.cliente_key, razon=S.razon,
                        nombre=f"{S.ap1}; {S.ap2}; {S.nom}",
                        sexo=S.sexo, edad=S.edad, ecivil=S.ecivil,
                        estudios=S.estudios, estatus=S.estatus,
                        puesto=S.puesto, area=S.area,
                        contrat=S.contrat, personal=S.personal,
                        jornada=S.jornada, rotacion=S.rotacion,
                        tpuesto=S.tpuesto, exp=S.exp,
                        respuestas=S.respuestas + ["Nunca","Nunca","Nunca","Nunca"],
                        resultado=resultado,
                    ))
                    S.pantalla = "fin"; st.rerun()
            st.stop()

    # ── Fin del cuestionario ──────────────────────────────────────────────────
    if idx >= len(preg_activas) and S.atiende_clientes is not None and S.es_jefe is not None:
        resultado = calcular_puntaje_g3(S.respuestas)
        S.res = resultado
        guardar(dict(
            folio=S.folio, cliente=S.cliente_key, razon=S.razon,
            nombre=f"{S.ap1}; {S.ap2}; {S.nom}",
            sexo=S.sexo, edad=S.edad, ecivil=S.ecivil,
            estudios=S.estudios, estatus=S.estatus,
            puesto=S.puesto, area=S.area,
            contrat=S.contrat, personal=S.personal,
            jornada=S.jornada, rotacion=S.rotacion,
            tpuesto=S.tpuesto, exp=S.exp,
            respuestas=S.respuestas,
            resultado=resultado,
        ))
        S.pantalla = "fin"; st.rerun()

    # ── Pregunta normal ───────────────────────────────────────────────────────
    preg = preg_activas[idx]
    sec_anterior = preg_activas[idx-1].get("sec","") if idx > 0 else ""
    mostrar_inst = preg.get("sec","") != sec_anterior

    if mostrar_inst and preg.get("sec","") in INSTRUCCIONES_G3:
        st.markdown(f"""
        <div style="background:#f0f6f0;border-left:4px solid #4b694e;border-radius:0 8px 8px 0;
                    padding:.8rem 1.2rem;margin-bottom:.8rem;font-size:.92rem;
                    color:#3a5a3e;font-style:italic;line-height:1.6;">
            {INSTRUCCIONES_G3[preg['sec']]}
        </div>
        """, unsafe_allow_html=True)

    st.markdown(
        f'<div class="prog-txt">{preg["cat_nombre"]} · '
        f'Pregunta {idx + 1} de {len(preg_activas)}</div>',
        unsafe_allow_html=True)
    st.progress(idx / max(len(preg_activas), 1))

    st.markdown(f"""
    <div class="pq-card">
      <div class="pq-num">Pregunta {idx + 1} / {len(preg_activas)}</div>
      <div class="pq-dom">Dominio: {preg["dom_nombre"]}</div>
      <div class="pq-txt">{preg["texto"]}</div>
    </div>
    """, unsafe_allow_html=True)

    resp = st.radio("", OPC_RESP, index=None, horizontal=False,
                    key=f"q_{idx}", label_visibility="collapsed")

    if S.err:
        st.markdown('<p class="err-r">⚠ SELECCIONA UNA RESPUESTA ANTES DE CONTINUAR</p>',
                    unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("SIGUIENTE →", use_container_width=True, key=f"sig_{idx}"):
        if resp is None:
            S.err = True; st.rerun()
        else:
            S.err        = False
            S.respuestas = S.respuestas + [resp]
            S.preg_idx   = idx + 1
            st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# FIN
# ══════════════════════════════════════════════════════════════════════════════
elif S.pantalla == "fin":
    header(False)
    st.markdown("""
    <div class="fin-box">
      <div class="fin-tit">✓ CUESTIONARIO FINALIZADO CORRECTAMENTE</div>
      <div class="fin-sub">AGRADECEMOS TU PARTICIPACIÓN</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    if _MODO_EMPLEADO:
        st.markdown("""
        <div style="text-align:center;font-size:.92rem;color:#666;
                    font-family:Montserrat,sans-serif;padding:1rem 0;">
            Puedes cerrar esta ventana.
        </div>
        """, unsafe_allow_html=True)
    else:
        if S.res and S.res.get("puntaje_total") is not None:
            r = S.res
            with st.expander("Ver resultado — uso interno operativo"):
                st.markdown(f"""
                <div style="background:{r.get('color','#4b694e')};color:#fff;border-radius:13px;
                            padding:1.5rem 2rem;text-align:center;margin-bottom:1rem;">
                    <div style="font-size:.75rem;opacity:.8;text-transform:uppercase;
                                letter-spacing:.1em;">Resultado · Guía III · NOM-035</div>
                    <div style="font-size:1.4rem;font-weight:700;margin-top:.4rem;">
                        {r.get('nivel','')}</div>
                    <div style="font-size:.9rem;opacity:.85;margin-top:.3rem;">
                        Puntaje total: {r.get('puntaje_total',0)} puntos</div>
                </div>
                """, unsafe_allow_html=True)
                if r.get("alerta_violencia"):
                    st.markdown("""
                    <div style="background:#fff0f0;border:2px solid #a20000;border-radius:11px;
                                padding:1rem 1.4rem;font-size:.9rem;color:#a20000;font-weight:700;">
                        🚨 ALERTA — VIOLENCIA LABORAL DETECTADA (ítems 57-64)
                    </div>
                    """, unsafe_allow_html=True)
        c1,c2 = st.columns(2)
        with c1:
            if st.button("REGISTRAR OTRO EMPLEADO", use_container_width=True):
                S.pantalla = "panel"; st.rerun()
        with c2:
            if st.button("VOLVER AL PANEL", use_container_width=True):
                S.pantalla = "panel"; st.rerun()
