import streamlit as st
import os
import random
from docx import Document
from pypdf import PdfReader

st.set_page_config(page_title="Coreano 3A - App de Estudio", page_icon="🇰🇷", layout="centered")

# --- BANCO COMPLETO Y EXTENSO DE VOCABULARIO (ÍNDICE PÁGS. 52-55) ---
BANCO_VOCABULARIO = [
    # 1, 3 (Pág. 52)
    {"coreano": "1등을 하다 / 1위를 하다", "espanol": "Ganar el primer lugar (Pág. 17)", "leccion": "Pág. 52"},
    {"coreano": "1주년", "espanol": "Primer aniversario (Pág. 14)", "leccion": "Pág. 52"},
    {"coreano": "100일", "espanol": "Día 100 (celebración) (Pág. 14)", "leccion": "Pág. 52"},
    {"coreano": "3·1절", "espanol": "Día del Movimiento de Independencia del 1 de Marzo (Pág. 16)", "leccion": "Pág. 52"},
    
    # ㄱ (Pág. 52)
    {"coreano": "광복절", "espanol": "Día de la Liberación (Pág. 16)", "leccion": "Pág. 52"},
    {"coreano": "교통이 편리하다", "espanol": "El transporte es conveniente (Pág. 8)", "leccion": "Pág. 52"},
    {"coreano": "교환하다", "espanol": "Cambiar / Intercambiar (Pág. 10)", "leccion": "Pág. 52"},
    {"coreano": "구입하다", "espanol": "Comprar / Adquirir (Pág. 10)", "leccion": "Pág. 52"},
    {"coreano": "국경일", "espanol": "Día festivo nacional (Pág. 16)", "leccion": "Pág. 52"},
    {"coreano": "그냥", "espanol": "Simplemente / Solo (Pág. 11)", "leccion": "Pág. 52"},
    {"coreano": "그저 그렇게 지내다", "espanol": "Pasar el tiempo sin nada especial (Pág. 6)", "leccion": "Pág. 52"},
    {"coreano": "근로자의 날", "espanol": "Día del Trabajo (Pág. 14)", "leccion": "Pág. 52"},
    {"coreano": "기념일을 맞이하다", "espanol": "Conmemorar un aniversario (Pág. 14)", "leccion": "Pág. 52"},
    {"coreano": "기념일을 챙기다", "espanol": "Celebrar un aniversario (Pág. 14)", "leccion": "Pág. 52"},
    {"coreano": "기념품을 받다", "espanol": "Recibir un souvenir (Pág. 17)", "leccion": "Pág. 52"},
    {"coreano": "기념품을 주다", "espanol": "Dar un souvenir (Pág. 17)", "leccion": "Pág. 52"},
    {"coreano": "기념행사를 하다", "espanol": "Hacer un evento conmemorativo (Pág. 14)", "leccion": "Pág. 52"},
    {"coreano": "기대에 못 미치다", "espanol": "No cumplir las expectativas (Pág. 22)", "leccion": "Pág. 52"},
    {"coreano": "기분이 가라앉다", "espanol": "Sentirse deprimido / Decaído (Pág. 18)", "leccion": "Pág. 52"},
    {"coreano": "기분 전환이 되다", "espanol": "Refrescar la mente / Cambiar de aire (Pág. 20)", "leccion": "Pág. 52"},
    {"coreano": "기침을 하다", "espanol": "Toser (Pág. 19)", "leccion": "Pág. 52"},
    {"coreano": "꽃을 달아 드리다", "espanol": "Regalar / Poner flores (Pág. 14)", "leccion": "Pág. 52"},

    # ㄴ, ㄷ (Pág. 53)
    {"coreano": "나이가 들다", "espanol": "Envejecer / Cumplir años (Pág. 22)", "leccion": "Pág. 53"},
    {"coreano": "날씨가 풀리다", "espanol": "El clima se vuelve más cálido (Pág. 22)", "leccion": "Pág. 53"},
    {"coreano": "남다", "espanol": "Quedar / Sobrar (Pág. 11)", "leccion": "Pág. 53"},
    {"coreano": "내용", "espanol": "Contenido (Pág. 9)", "leccion": "Pág. 53"},
    {"coreano": "눈물이 나다", "espanol": "Llorar / Salir lágrimas (Pág. 18)", "leccion": "Pág. 53"},
    {"coreano": "다리를 떨다", "espanol": "Mover / Temblar la pierna por nervios (Pág. 24)", "leccion": "Pág. 53"},
    {"coreano": "답답하다", "espanol": "Sentirse sofocado / Frustrado (Pág. 18)", "leccion": "Pág. 53"},
    {"coreano": "대단하다", "espanol": "Ser increíble / Grandioso (Pág. 17)", "leccion": "Pág. 53"},
    {"coreano": "대화하다", "espanol": "Conversar / Dialogar (Pág. 6)", "leccion": "Pág. 53"},
    {"coreano": "더위 / 추위", "espanol": "El calor / El frío (Pág. 22)", "leccion": "Pág. 53"},
    {"coreano": "도움이 되다", "espanol": "Ser de ayuda (Pág. 20)", "leccion": "Pág. 53"},
    {"coreano": "동창회", "espanol": "Reunión de exalumnos (Pág. 12)", "leccion": "Pág. 53"},

    # ㅁ, ㅂ (Pág. 54)
    {"coreano": "마음에 들다", "espanol": "Ser del agrado de uno (Pág. 10)", "leccion": "Pág. 54"},
    {"coreano": "마음을 놓다", "espanol": "Tranquilizarse / Relajarse (Pág. 20)", "leccion": "Pág. 54"},
    {"coreano": "만족스럽다", "espanol": "Estar satisfecho (Pág. 22)", "leccion": "Pág. 54"},
    {"coreano": "말씀을 드리다", "espanol": "Decir / Hablar (forma honorífica) (Pág. 6)", "leccion": "Pág. 54"},
    {"coreano": "몸살이 나다", "espanol": "Tener dolor muscular por fatiga/gripe (Pág. 19)", "leccion": "Pág. 54"},
    {"coreano": "무기력하다", "espanol": "Estar sin energía / Letárgico (Pág. 18)", "leccion": "Pág. 54"},
    {"coreano": "바라다", "espanol": "Desear / Esperar algo (Pág. 17)", "leccion": "Pág. 54"},
    {"coreano": "바꾸다", "espanol": "Cambiar / Intercambiar (Pág. 10)", "leccion": "Pág. 54"},
    {"coreano": "방문하다", "espanol": "Visitar (Pág. 12)", "leccion": "Pág. 54"},
    {"coreano": "버리다", "espanol": "Tirar / Desechar (Pág. 8)", "leccion": "Pág. 54"},
    {"coreano": "불편하다", "espanol": "Incómodo / Inconveniente (Pág. 8)", "leccion": "Pág. 54"},
    {"coreano": "비용이 들다", "espanol": "Tener un costo / Gasto (Pág. 10)", "leccion": "Pág. 54"},

    # ㅅ, ㅇ, ㅈ (Pág. 55)
    {"coreano": "사고가 나다", "espanol": "Ocurrir un accidente (Pág. 21)", "leccion": "Pág. 55"},
    {"coreano": "새로 사다", "espanol": "Comprar nuevo (Pág. 8)", "leccion": "Pág. 55"},
    {"coreano": "설거지를 하다", "espanol": "Lavar los platos (Pág. 8)", "leccion": "Pág. 55"},
    {"coreano": "성년의 날", "espanol": "Día de la Mayoría de Edad (Pág. 14)", "leccion": "Pág. 55"},
    {"coreano": "수리하다", "espanol": "Reparar / Arreglar (Pág. 8)", "leccion": "Pág. 55"},
    {"coreano": "스승의 날", "espanol": "Día del Maestro (Pág. 14)", "leccion": "Pág. 55"},
    {"coreano": "신나다", "espanol": "Estar emocionado / Animado (Pág. 18)", "leccion": "Pág. 55"},
    {"coreano": "약속을 잡다", "espanol": "Fijar una cita / Compromiso (Pág. 12)", "leccion": "Pág. 55"},
    {"coreano": "어린이날", "espanol": "Día del Niño (Pág. 14)", "leccion": "Pág. 55"},
    {"coreano": "어버이날", "espanol": "Día de los Padres (Pág. 14)", "leccion": "Pág. 55"},
    {"coreano": "열이 나다", "espanol": "Tener fiebre (Pág. 19)", "leccion": "Pág. 55"},
    {"coreano": "위로하다", "espanol": "Consolar / Dar ánimos (Pág. 20)", "leccion": "Pág. 55"},
    {"coreano": "응원하다", "espanol": "Apoyar / Animar (Pág. 17)", "leccion": "Pág. 55"},
    {"coreano": "자리를 비우다", "espanol": "Ausentarse del asiento (Pág. 24)", "leccion": "Pág. 55"},
    {"coreano": "자랑스럽다", "espanol": "Estar orgulloso (Pág. 18)", "leccion": "Pág. 55"},
    {"coreano": "축하하다", "espanol": "Felicitar / Celebrar (Pág. 14)", "leccion": "Pág. 55"},
    {"coreano": "한글날", "espanol": "Día del Hangul (Pág. 16)", "leccion": "Pág. 55"},
    {"coreano": "환불하다", "espanol": "Reembolsar / Reembolso (Pág. 10)", "leccion": "Pág. 55"}
]

# --- CARGAR APUNTES WORD ---
@st.cache_data
def cargar_gramatica_word():
    gramatica_dict = {}
    for archivo in os.listdir("."):
        if archivo.endswith(".docx"):
            try:
                doc = Document(archivo)
                texto = "\n".join([p.text for p in doc.paragraphs if p.text.strip() != ""])
                gramatica_dict[f"Apunte: {archivo.replace('.docx', '')}"] = texto
            except Exception:
                pass
    return gramatica_dict

BANCO_GRAMATICA = [
    {
        "pregunta": "¿Qué expresa la estructura '-을 테니까 / ㄹ 테니까'?",
        "opciones": [
            "Una suposición o promesa propia, sugiriendo o pidiendo algo al oyente.",
            "Una obligación legal irrenunciable.",
            "Una acción del pasado totalmente olvidada.",
            "Una disculpa por llegar tarde a una cita."
        ],
        "correcta": "Una suposición o promesa propia, sugiriendo o pidiendo algo al oyente."
    },
    {
        "pregunta": "¿Cuándo se utiliza la terminación '-는 데 / 은 데'?",
        "opciones": [
            "Para proporcionar información de fondo o contexto antes de sugerir o preguntar algo.",
            "Para indicar el final definitivo de una oración sin dar explicaciones.",
            "Para contar números cardinales en coreano puro.",
            "Para expresar prohibición absoluta."
        ],
        "correcta": "Para proporcionar información de fondo o contexto antes de sugerir o preguntar algo."
    }
]

st.title("🇰🇷 Mi App de Coreano 3A")
st.caption(f"Vocabulario total cargado: {len(BANCO_VOCABULARIO)} palabras del Índice (Págs. 52-55)")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🎴 Vocabulario", 
    "✍️ Examen Vocabulario", 
    "📖 Gramática (Notas)", 
    "📝 Examen Gramática",
    "📚 Libro PDF"
])

# -------------------------------------------------------------
# TAB 1: TARJETAS DE VOCABULARIO
# -------------------------------------------------------------
with tab1:
    st.header("🎴 Tarjetas de Repaso (Índice Completo)")
    if "flash_idx" not in st.session_state:
        st.session_state.flash_idx = 0
    if "mostrar_trad" not in st.session_state:
        st.session_state.mostrar_trad = False
        
    card = BANCO_VOCABULARIO[st.session_state.flash_idx % len(BANCO_VOCABULARIO)]
    
    st.info(f"📍 {card['leccion']} | Tarjeta {st.session_state.flash_idx + 1} de {len(BANCO_VOCABULARIO)}")
    st.markdown(f"<h1 style='text-align: center; color: #2E86C1; margin: 30px 0;'>{card['coreano']}</h1>", unsafe_allow_html=True)
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("👁️ Ver Traducción", use_container_width=True):
            st.session_state.mostrar_trad = True
    with c2:
        if st.button("➡️ Siguiente Palabra", use_container_width=True):
            st.session_state.flash_idx += 1
            st.session_state.mostrar_trad = False
            st.rerun()
            
    if st.session_state.mostrar_trad:
        st.success(f"### Significado: {card['espanol']}")

# -------------------------------------------------------------
# TAB 2: EXAMEN DE VOCABULARIO
# -------------------------------------------------------------
with tab2:
    st.header("✍️ Examen de Vocabulario")
    if "quiz_idx" not in st.session_state:
        st.session_state.quiz_idx = 0
        st.session_state.evaluado_v = False
        st.session_state.res_v = False

    p_act = BANCO_VOCABULARIO[st.session_state.quiz_idx % len(BANCO_VOCABULARIO)]
    st.subheader(f"¿Cuál es el significado de: **{p_act['coreano']}**?")
    
    otras_opciones = [v["espanol"] for v in BANCO_VOCABULARIO if v["espanol"] != p_act["espanol"]]
    ops = random.sample(otras_opciones, min(3, len(otras_opciones))) + [p_act["espanol"]]
    random.shuffle(ops)
    
    sel_v = st.radio("Selecciona la respuesta correcta:", ops, key=f"v_{st.session_state.quiz_idx}")
    
    ac1, ac2 = st.columns(2)
    with ac1:
        if not st.session_state.evaluado_v:
            if st.button("Comprobar Respuesta", use_container_width=True):
                st.session_state.evaluado_v = True
                st.session_state.res_v = (sel_v == p_act["espanol"])
                st.rerun()
        else:
            if st.session_state.res_v:
                st.success("¡Excelente! 🎉 Respuesta correcta.")
            else:
                st.error(f"Incorrecto ❌. La traducción correcta es: **{p_act['espanol']}**")
    with ac2:
        if st.session_state.evaluado_v:
            if st.button("Siguiente Pregunta ➡️", use_container_width=True):
                st.session_state.quiz_idx += 1
                st.session_state.evaluado_v = False
                st.rerun()

# -------------------------------------------------------------
# TAB 3: GRAMÁTICA (DOCX)
# -------------------------------------------------------------
with tab3:
    st.header("📖 Apuntes de Gramática")
    gram_todo = cargar_gramatica_word()
    
    if gram_todo:
        f_sel = st.selectbox("Selecciona un documento:", list(gram_todo.keys()))
        st.markdown("---")
        st.text_area("Contenido:", gram_todo[f_sel], height=400)
    else:
        st.info("Agrega archivos `.docx` en tu carpeta de proyecto para verlos aquí.")

# -------------------------------------------------------------
# TAB 4: EXAMEN DE GRAMÁTICA
# -------------------------------------------------------------
with tab4:
    st.header("📝 Examen de Gramática")
    if "g_idx" not in st.session_state:
        st.session_state.g_idx = 0
        st.session_state.eval_g = False
        st.session_state.res_g = False

    pregunta_g = BANCO_GRAMATICA[st.session_state.g_idx % len(BANCO_GRAMATICA)]
    st.subheader(pregunta_g["pregunta"])
    
    opciones_g = pregunta_g["opciones"].copy()
    
    sel_g = st.radio("Selecciona una opción:", opciones_g, key=f"g_radio_{st.session_state.g_idx}")
    
    gc1, gc2 = st.columns(2)
    with gc1:
        if not st.session_state.eval_g:
            if st.button("Evaluar Gramática", use_container_width=True):
                st.session_state.eval_g = True
                st.session_state.res_g = (sel_g == pregunta_g["correcta"])
                st.rerun()
        else:
            if st.session_state.res_g:
                st.success("¡Respuesta Correcta! 🎯")
            else:
                st.error(f"Incorrecto ❌. Opción correcta:\n\n**{pregunta_g['correcta']}**")
    with gc2:
        if st.session_state.eval_g:
            if st.button("Siguiente Ejercicio ➡️", use_container_width=True):
                st.session_state.g_idx += 1
                st.session_state.eval_g = False
                st.rerun()

# -------------------------------------------------------------
# TAB 5: LIBRO PDF COMPLETO
# -------------------------------------------------------------
with tab5:
    st.header("📚 Lector de PDF del Libro")
    archivo_pdf = "Libro Coreno 3A.pdf"
    if os.path.exists(archivo_pdf):
        lector = PdfReader(archivo_pdf)
        num_p = st.slider("Selecciona el número de página:", 1, len(lector.pages), 52)
        st.text_area("Contenido de la página:", lector.pages[num_p-1].extract_text(), height=400)
    else:
        st.warning("No se encontró el archivo del libro PDF en el directorio del proyecto.")