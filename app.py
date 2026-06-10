import streamlit as st
import random
import string
import unicodedata
import io
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Generador de Sopa de Letras",
    page_icon="🧩",
    layout="centered"
)

# ==========================================
# FUNCIONES DEL ALGORITMO
# ==========================================
def limpiar_texto(texto):
    """
    Limpia el texto: quita espacios, elimina tildes pero conserva la 'Ñ',
    y pasa todo a mayúsculas.
    """
    texto = texto.upper().strip()
    texto = texto.replace(" ", "")
    
    # Proteger la Ñ antes de quitar tildes
    texto = texto.replace('Ñ', '@N@')
    
    # Quitar tildes (normalización NFD)
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    
    # Restaurar la Ñ
    texto = texto.replace('@N@', 'Ñ')
    
    # Dejar solo letras (A-Z y Ñ)
    return ''.join(c for c in texto if c.isalpha() or c == 'Ñ')

def generar_sopa(palabras_originales):
    """
    Crea la matriz de la sopa de letras y la matriz de la solución.
    """
    # Limpiar y filtrar palabras válidas
    mapeo_palabras = []
    for p in palabras_originales:
        limpia = limpiar_texto(p)
        if limpia:
            mapeo_palabras.append((p.strip(), limpia))
            
    if not mapeo_palabras:
        return None, None, []

    palabras_limpias = [p[1] for p in mapeo_palabras]
    
    # 1. Calcular tamaño dinámico de la cuadrícula
    largo_max = max(len(p) for p in palabras_limpias)
    total_letras = sum(len(p) for p in palabras_limpias)
    
    # Fórmula empírica para que haya espacio suficiente para cruzarse
    n = max(largo_max + 2, int(total_letras ** 0.5 * 1.5))
    n = min(max(n, 12), 25) # Mantener entre 12x12 y 25x25 para que entre en el Word

    # Inicializar cuadrículas vacías
    grid = [['' for _ in range(n)] for _ in range(n)]
    solucion = [[' ' for _ in range(n)] for _ in range(n)]

    # Todas las direcciones (Horizontal, Vertical, Diagonales y sus inversas)
    direcciones = [(0, 1), (1, 0), (1, 1), (-1, 1), (0, -1), (-1, 0), (-1, -1), (1, -1)]

    palabras_colocadas = []

    # 2. Posicionar palabras
    for original, limpia in mapeo_palabras:
        colocada = False
        intentos = 0
        
        while not colocada and intentos < 300:
            df, dc = random.choice(direcciones)
            f = random.randint(0, n - 1)
            c = random.randint(0, n - 1)
            
            # Verificar si la palabra cabe en los límites
            if 0 <= f + df * (len(limpia) - 1) < n and 0 <= c + dc * (len(limpia) - 1) < n:
                cabe = True
                # Verificar colisiones
                for i, letra in enumerate(limpia):
                    if grid[f + df * i][c + dc * i] not in ('', letra):
                        cabe = False
                        break
                
                # Si cabe y no choca, se coloca
                if cabe:
                    for i, letra in enumerate(limpia):
                        grid[f + df * i][c + dc * i] = letra
                        solucion[f + df * i][c + dc * i] = letra
                    colocada = True
                    palabras_colocadas.append(original)
            intentos += 1

    # 3. Rellenar espacios vacíos con letras aleatorias
    letras_abecedario = string.ascii_uppercase.replace('Q', 'O').replace('W', 'A') + 'Ñ' # Ajuste para español
    for i in range(n):
        for j in range(n):
            if grid[i][j] == '':
                grid[i][j] = random.choice(letras_abecedario)

    return grid, solucion, palabras_colocadas

# ==========================================
# EXPORTACIÓN A WORD
# ==========================================
def exportar_a_word(grid, solucion, palabras):
    """
    Genera el archivo .docx con el formato exacto requerido.
    """
    doc = Document()
    
    # Ajustar márgenes para maximizar espacio
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # ================= PÁGINA 1: EL JUEGO =================
    titulo1 = doc.add_paragraph()
    titulo1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t1 = titulo1.add_run("SOPA DE LETRAS")
    run_t1.font.size = Pt(20)
    run_t1.bold = True
    
    doc.add_paragraph() # Espaciador

    # Cuadrícula del juego
    p_grid = doc.add_paragraph()
    p_grid.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_grid.paragraph_format.line_spacing = 1.2 # Interlineado para matriz cuadrada
    
    for fila in grid:
        texto_fila = "  ".join(fila) # Dos espacios entre letras
        run_fila = p_grid.add_run(texto_fila + "\n")
        run_fila.font.name = 'Consolas' # Fuente monoespaciada obligatoria
        run_fila.font.size = Pt(14)

    doc.add_paragraph()
    
    # Lista de palabras a buscar
    p_sub = doc.add_paragraph()
    run_sub = p_sub.add_run("PALABRAS A BUSCAR:")
    run_sub.bold = True
    
    for p in sorted(palabras):
        doc.add_paragraph(f"• {p.upper()}", style='List Bullet')

    # ================= PÁGINA 2: LA SOLUCIÓN =================
    doc.add_page_break() # Salto de página automático

    titulo2 = doc.add_paragraph()
    titulo2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t2 = titulo2.add_run("SOLUCIÓN")
    run_t2.font.size = Pt(20)
    run_t2.bold = True
    
    doc.add_paragraph()

    # Cuadrícula de la solución
    p_sol = doc.add_paragraph()
    p_sol.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sol.paragraph_format.line_spacing = 1.2
    
    for fila in solucion:
        # Aquí ya hay espacios en blanco donde no hay palabras
        texto_fila = "  ".join(fila) 
        run_fsol = p_sol.add_run(texto_fila + "\n")
        run_fsol.font.name = 'Consolas'
        run_fsol.font.size = Pt(14)
        run_fsol.bold = True

    # Guardar en memoria
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# ==========================================
# INTERFAZ DE USUARIO (UI)
# ==========================================
st.title("🧩 Generador de Sopa de Letras")
st.markdown("Ingresa tu lista de palabras. El sistema creará una sopa de letras perfecta para imprimir en Word, con las soluciones en la segunda página.")

# Inicializar variables de sesión para no perder datos al descargar
if 'docx_bytes' not in st.session_state:
    st.session_state.docx_bytes = None
if 'exito' not in st.session_state:
    st.session_state.exito = False
if 'total_palabras' not in st.session_state:
    st.session_state.total_palabras = 0

# Palabras precargadas para probar (separadas por salto de línea)
default_words = "IPERC\nLOTO\nKPREVENCION\nSALUD OCUPACIONAL\n29783\nINSPECCION\nRIESGOS\nEPP"

texto_ingresado = st.text_area(
    "Palabras o frases (Una por línea, usa 'Enter' para separar):", 
    value=default_words,
    height=200
)

# Botón principal
if st.button("🚀 Generar Sopa de Letras", type="primary"):
    lineas = texto_ingresado.split('\n')
    lineas = [l for l in lineas if l.strip()] # Quitar líneas vacías
    
    if len(lineas) > 0:
        with st.spinner("Generando cuadrícula y cruzando palabras..."):
            grid, solucion, colocadas = generar_sopa(lineas)
            
            if grid:
                # Generar Word y guardar en session_state
                docx_bytes = exportar_a_word(grid, solucion, colocadas)
                st.session_state.docx_bytes = docx_bytes
                st.session_state.total_palabras = len(colocadas)
                st.session_state.exito = True
            else:
                st.error("No se detectaron palabras válidas con caracteres alfabéticos.")
    else:
        st.warning("Por favor, ingresa al menos una palabra.")

# Si hay un documento en memoria, mostrar botón de descarga
if st.session_state.exito and st.session_state.docx_bytes:
    st.success(f"¡Sopa de letras generada con éxito! Se escondieron {st.session_state.total_palabras} palabras.")
    
    st.download_button(
        label="📄 Descargar Documento Word (.docx)",
        data=st.session_state.docx_bytes,
        file_name="Sopa_de_Letras_Generada.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
