import streamlit as st
import random
import string
import io
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Configuración de la página
st.set_page_config(page_title="Generador de Sopa de Letras", page_icon="🧩", layout="centered")

def create_word_search(words):
    # Limpiar palabras (quitar espacios y pasar a mayúsculas)
    clean_words = [w.replace(" ", "").upper() for w in words if w.strip()]
    if not clean_words:
        return None, None, None

    # Calcular un tamaño de cuadrícula dinámico y seguro
    max_len = max(len(w) for w in clean_words)
    total_chars = sum(len(w) for w in clean_words)
    grid_size = max(15, max_len + 2, int(total_chars**0.5 * 1.8))
    
    grid = [['' for _ in range(grid_size)] for _ in range(grid_size)]
    solution = [[' ' for _ in range(grid_size)] for _ in range(grid_size)]
    
    # Direcciones: Horizontal, Vertical, Diagonal (ambos sentidos)
    directions = [(0,1), (1,0), (1,1), (-1,1), (0,-1), (-1,0), (-1,-1), (1,-1)]
    
    placed_words = []
    
    for original_word, word in zip([w for w in words if w.strip()], clean_words):
        placed = False
        attempts = 0
        while not placed and attempts < 300:
            dr, dc = random.choice(directions)
            r = random.randint(0, grid_size - 1)
            c = random.randint(0, grid_size - 1)
            
            # Verificar si la palabra cabe en esa dirección
            if 0 <= r + dr*(len(word)-1) < grid_size and 0 <= c + dc*(len(word)-1) < grid_size:
                match = True
                for i, letter in enumerate(word):
                    nr, nc = r + dr*i, c + dc*i
                    if grid[nr][nc] != '' and grid[nr][nc] != letter:
                        match = False
                        break
                
                # Si cabe y no choca, colocarla
                if match:
                    for i, letter in enumerate(word):
                        nr, nc = r + dr*i, c + dc*i
                        grid[nr][nc] = letter
                        solution[nr][nc] = letter # Marcar en la solución
                    placed = True
                    placed_words.append(original_word.strip().upper())
            attempts += 1
            
    # Rellenar los espacios vacíos con letras aleatorias
    for r in range(grid_size):
        for c in range(grid_size):
            if grid[r][c] == '':
                grid[r][c] = random.choice(string.ascii_uppercase)
            if solution[r][c] == ' ':
                solution[r][c] = '.' # Puntos para que la solución sea clara

    return grid, solution, placed_words

def generate_word_doc(grid, solution, original_words):
    doc = Document()
    
    # Configurar estilo para que sea monoespaciado y cuadre perfecto
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Courier New'
    font.size = Pt(14)
    
    # --- PÁGINA 1: LA SOPA DE LETRAS ---
    title = doc.add_heading('SOPA DE LETRAS', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph() # Espacio
    
    grid_para = doc.add_paragraph()
    grid_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for row in grid:
        grid_para.add_run('  '.join(row) + '\n')
        
    doc.add_paragraph() # Espacio
    doc.add_heading('Palabras a buscar:', level=2)
    for w in original_words:
        doc.add_paragraph(w, style='List Bullet')
        
    # --- PÁGINA 2: LA SOLUCIÓN ---
    doc.add_page_break()
    
    sol_title = doc.add_heading('SOLUCIÓN', 0)
    sol_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph() # Espacio
    
    sol_para = doc.add_paragraph()
    sol_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for row in solution:
        # Reemplazar los puntos por espacios para que solo resalten las palabras
        row_display = [char if char != '.' else ' ' for char in row]
        sol_para.add_run('  '.join(row_display) + '\n')

    # Guardar en memoria para descargar
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- INTERFAZ DE STREAMLIT ---
st.title("🧩 Generador de Sopa de Letras")
st.markdown("Ingresa una palabra o frase por línea. El sistema ignorará los espacios al esconder las palabras en la cuadrícula.")

# Palabras por defecto
default_words = "IPERC\nLOTO\nPREVENCION\nSEGURIDAD\nSALUD OCUPACIONAL\nINSPECCION\nRIESGOS\nEPP"

text_input = st.text_area("Lista de palabras (una por línea):", value=default_words, height=200)

if st.button("🚀 Generar Sopa de Letras", type="primary"):
    words_list = text_input.split('\n')
    
    with st.spinner('Generando la cuadrícula y cruzando palabras...'):
        grid, solution, placed_words = create_word_search(words_list)
        
        if grid:
            st.success(f"¡Listo! Se logró colocar {len(placed_words)} palabras.")
            
            # Generar el archivo Word
            docx_file = generate_word_doc(grid, solution, placed_words)
            
            # Botón de descarga
            st.download_button(
                label="📄 Descargar en Word",
                data=docx_file,
                file_name="sopa_de_letras.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            # Vista previa rápida en la web
            st.markdown("### Vista Previa")
            preview_cols = st.columns(2)
            with preview_cols[0]:
                st.caption("Sopa de Letras")
                st.code('\n'.join([' '.join(row) for row in grid]))
            with preview_cols[1]:
                st.caption("Solución")
                st.code('\n'.join([' '.join([c if c != '.' else ' ' for c in row]) for row in solution]))
        else:
            st.error("No se detectaron palabras válidas. Por favor ingresa texto.")