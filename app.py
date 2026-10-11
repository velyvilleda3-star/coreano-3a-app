import streamlit as st
import os
import random
import re
from docx import Document
from pypdf import PdfReader

st.set_page_config(page_title="Coreano 3A - App de Estudio", page_icon="🇰🇷", layout="centered")

# --- EXTRAER TODO EL VOCABULARIO (Pág. 8 a 24 del PDF) ---
@st.cache_data
def extraer_todo_vocabulario_pdf():
    archivo_pdf = "Libro Coreno 3A.pdf"
    lista_vocab = []
    
    # Vocabulario de respaldo estructurado
    vocab_base = [
        {"coreano": "청소를 하다", "espanol": "Hacer la limpieza", "leccion": "Pág. 8-24"},
        {"coreano": "주방", "espanol": "Cocina", "leccion": "Pág. 8-24"},
        {"coreano": "거실", "espanol": "Sala de estar", "leccion": "Pág. 8-24"},
        {"coreano": "설거지를 하다", "espanol": "Lavar los platos", "leccion": "Pág. 8-24"},
        {"coreano": "빨래를 하다", "espanol": "Lavar la ropa", "leccion": "Pág. 8-24"},
        {"coreano": "수리하다", "espanol": "Reparar / Arreglar", "leccion": "Pág. 8-24"},
        {"coreano": "새로 사다", "espanol": "Comprar nuevo", "leccion": "Pág. 8-24"},
        {"coreano": "고장 나다", "espanol": "Descomponerse / Romperse", "leccion": "Pág. 8-24"},
        {"coreano": "버리다", "espanol": "Tirar / Desechar", "leccion": "Pág. 8-24"},
        {"coreano": "바꾸다", "espanol": "Cambiar / Cambiarse", "leccion": "Pág. 8-24"}
    ]
    
    if os.path.exists(archivo_pdf):
        try:
            lector = PdfReader(archivo_pdf)
            total_pags = len(lector.pages)
            
            # Recorrer páginas 8 a 24
            fin_pag = min(24, total_pags)
            for i in range(7, fin_pag):
                texto = lector.pages[i].extract_text()
                if texto:
                    lineas = texto.split("\n")
                    for linea in lineas:
                        # Detectar caracteres en coreano (Hangul)
                        coreano = "".join(re.findall(r'[\uac00-\ud7a3]+', linea))
                        if len(coreano) > 1:
                            lista_vocab.append({
                                "coreano": linea.strip(),
                                "espanol": f"Vocabulario extraído de la Pág. {i+1}",
                                "leccion": f"Página {i+1}"
                            })
        except Exception:
            pass

    return lista_vocab if len(lista_vocab) > 5 else vocab_base

# --- CARGAR GRAMÁTICA (Word y Pág. 25+ PDF) ---
@st.cache_data
def cargar_gramatica_word_y_pdf():
    gramatica_dict = {}
    
    # Archivos Word (.docx)
    for archivo in os.listdir("."):
        if archivo.endswith(".docx"):
            try:
                doc = Document(archivo)
                texto = "\n".join([p.text for p in doc.paragraphs if p.text.strip() != ""])
                nombre = f"Word: {archivo.replace('.docx', '')}"
                gramatica_dict[nombre] = texto
            except Exception:
                pass
                
    # Libro PDF (Pág 25 en adelante)
    archivo_pdf = "Libro Coreno 3A.pdf"
    if os.path.exists(archivo_pdf):
        try:
            lector = PdfReader(archivo_pdf)
            for i in range(24, len(lector.pages)):
                t = lector.pages[i].extract_text()
                if t and len(t.strip()) > 80:
                    gramatica_dict[f"Libro PDF - Pág. {i+1}"] = t
        except Exception:
            pass
            
    return gramatica_dict

BANCO_VOCABULARIO = extraer_todo_vocabulario_pdf()

BANCO_GRAMATICA_AVANZADO = [
    {
        "pregunta": "¿Qué expresa principalmente la estructura '-을 테니까 / ㄹ 테니까'?",
        "opciones": [
            "Una suposición o promesa propia, esperando que el oyente actúe en consecuencia.",
            "Una obligación legal obligatoria impuesta por el gobierno.",
            "Una queja sobre un evento ocurrido en la infancia.",
            "Una orden directa para limpiar la cocina inmediatamente."
        ],
        "correcta": "Una suposición o promesa propia, esperando que el oyente actúe en consecuencia."
    },
    {
        "pregunta": "¿Cuál es el contexto de uso de la terminación conectiva de fondo '-는 데 / 은 데'?",
        "opciones": [
            "Presentar un contexto previo o situación antecedente antes de hacer una sugerencia o pregunta.",
            "Indicar una prohibición estricta de paso en la vía pública.",
            "Describir exclusivamente objetos que ya no funcionan.",
            "Señalar el tiempo futuro lejano sin excepciones."
        ],
        "correcta": "Presentar un contexto previo o situación antecedente antes de hacer una sugerencia o pregunta."
    },
    {
        "pregunta": "¿Cuándo se utiliza la expresión de suposición subjetiva '-는 것 같다'?",
        "opciones": [
            "Cuando se expresa una impresión o parecer personal basado en lo que se observa o experimenta.",
            "Cuando se quiere dar una orden imperativa de carácter militar.",
            "Cuando se describe un hecho científico totalmente comprobado.",
            "Cuando se niega de forma categórica un suceso pasado."
        ],
        "correcta": "Cuando se expresa una impresión o parecer personal basado en lo que se observa o experimenta."
    }
]

st.title("🇰🇷 Mi Aplicación de Coreano 3A")
st.caption(f"Vocabulario completo cargado ({len(BANCO_VOCABULARIO)} elementos) + Gramática del libro y notas")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🎴 Vocabulario (Pág. 8-24)", 
    "✍️ Examen Vocabulario", 
    "📖 Gramática (Libro y Word)", 
    "📝 Examen de Gramática",
    "📚 Libro PDF"
])

# -------------------------------------------------------------
# PESTAÑA 1: VOCABULARIO COMPLETO
# -------------------------------------------------------------
with tab1:
    st.header("🎴 Tarjetas de Vocabulario Completo")
    if "flash_idx" not in st.session_state:
        st.session_state.flash_idx = 0
    if "mostrar_trad" not in st.session_state:
        st.session_state.mostrar_trad = False
        
    card = BANCO_VOCABULARIO[st.session_state.flash_idx % len(BANCO_VOCABULARIO)]
    st.info(f"📚 {card['leccion']}")
    st.markdown(f"<h1 style='text-align: center; color: #2E86C1;'>{card['coreano']}</h1>", unsafe_allow_html=True)
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("👁️ Ver Significado", use_container_width=True):
            st.session_state.mostrar_trad = True
    with c2:
        if st.button("➡️ Siguiente Tarjeta", use_container_width=True):
            st.session_state.flash_idx += 1
            st.session_state.mostrar_trad = False
            st.rerun()
            
    if st.session_state.mostrar_trad:
        st.success(f"### Significado: {card['espanol']}")

# -------------------------------------------------------------
# PESTAÑA 2: EXAMEN DE VOCABULARIO
# -------------------------------------------------------------
with tab2:
    st.header("✍️ Examen de Vocabulario Completo")
    if "quiz_idx" not in st.session_state:
        st.session_state.quiz_idx = 0
        st.session_state.puntuacion_v = 0
        st.session_state.evaluado_v = False
        st.session_state.res_v = False

    p_act = BANCO_VOCABULARIO[st.session_state.quiz_idx % len(BANCO_VOCABULARIO)]
    st.subheader(f"¿Qué significa **{p_act['coreano']}**?")
    
    inc = [v["espanol"] for v in BANCO_VOCABULARIO if v["espanol"] != p_act["espanol"]]
    ops = random.sample(inc, min(3, len(inc))) + [p_act["espanol"]]
    random.shuffle(ops)
    
    sel_v = st.radio("Opciones:", ops, key=f"v_{st.session_state.quiz_idx}")
    
    ac1, ac2 = st.columns(2)
    with ac1:
        if not st.session_state.evaluado_v:
            if st.button("Comprobar", use_container_width=True):
                st.session_state.evaluado_v = True
                if sel_v == p_act["espanol"]:
                    st.session_state.res_v = True
                    st.session_state.puntuacion_v += 1
                else:
                    st.session_state.res_v = False
                st.rerun()
        else:
            if st.session_state.res_v:
                st.success("¡Correcto! 🎉")
            else:
                st.error(f"Incorrecto ❌. Era: **{p_act['espanol']}**")
    with ac2:
        if st.session_state.evaluado_v:
            if st.button("Siguiente ➡️", use_container_width=True):
                st.session_state.quiz_idx += 1
                st.session_state.evaluado_v = False
                st.rerun()

# -------------------------------------------------------------
# PESTAÑA 3: GRAMÁTICA
# -------------------------------------------------------------
with tab3:
    st.header("📖 Repaso de Gramática (Libro y Apuntes)")
    gram_todo = cargar_gramatica_word_y_pdf()
    
    if gram_todo:
        f_sel = st.selectbox("Selecciona la lección:", list(gram_todo.keys()))
        st.markdown(f"### 📝 {f_sel}")
        st.markdown("---")
        st.text_area("Contenido:", gram_todo[f_sel], height=400)
    else:
        st.warning("No se encontraron contenidos de gramática.")

# -------------------------------------------------------------
# PESTAÑA 4: EXAMEN DE GRAMÁTICA
# -------------------------------------------------------------
with tab4:
    st.header("📝 Examen Interactivo de Gramática")
    
    if "g_idx" not in st.session_state:
        st.session_state.g_idx = 0
        st.session_state.punt_g = 0
        st.session_state.eval_g = False
        st.session_state.res_g = False

    pregunta_g = BANCO_GRAMATICA_AVANZADO[st.session_state.g_idx % len(BANCO_GRAMATICA_AVANZADO)]
    
    st.subheader(pregunta_g["pregunta"])
    
    opciones_g = pregunta_g["opciones"].copy()
    random.shuffle(opciones_g)
    
    sel_g = st.radio("Selecciona la respuesta:", opciones_g, key=f"g_radio_{st.session_state.g_idx}")
    
    gc1, gc2 = st.columns(2)
    with gc1:
        if not st.session_state.eval_g:
            if st.button("Comprobar Gramática", use_container_width=True):
                st.session_state.eval_g = True
                if sel_g == pregunta_g["correcta"]:
                    st.session_state.res_g = True
                    st.session_state.punt_g += 1
                else:
                    st.session_state.res_g = False
                st.rerun()
        else:
            if st.session_state.res_g:
                st.success("¡Excelente! Respuesta correcta 🎯")
            else:
                st.error(f"Incorrecto ❌. La respuesta correcta era:\n\n**{pregunta_g['correcta']}**")
    with gc2:
        if st.session_state.eval_g:
            if st.button("Siguiente Gramática ➡️", use_container_width=True):
                st.session_state.g_idx += 1
                st.session_state.eval_g = False
                st.rerun()

# -------------------------------------------------------------
# PESTAÑA 5: LIBRO PDF COMPLETO
# -------------------------------------------------------------
with tab5:
    st.header("📚 Libro PDF Completo")
    archivo_pdf = "Libro Coreno 3A.pdf"
    if os.path.exists(archivo_pdf):
        lector = PdfReader(archivo_pdf)
        num_p = st.slider("Página del libro:", 1, len(lector.pages), 1)
        st.text_area("Texto de la página:", lector.pages[num_p-1].extract_text(), height=400)
    else:
        st.warning("No se encontró el PDF.")