import streamlit as st
import os
import re
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
from io import BytesIO

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="KKP Generator AI", layout="wide")

AVAILABLE_MODELS = [
    "gemini-1.5-flash", "gemini-1.5-pro", "gemini-2.5-flash", 
    "gemini-3-flash-preview", "gemini-3-pro-preview", "Input Manual..."
]

# --- SYSTEM INSTRUCTION (OPTIMAL FOR ALIGNMENT) ---
SYSTEM_INSTRUCTION = """
Anda adalah Auditor Senior SPI PT AGRINAS PANGAN NUSANTARA (PERSERO).
Susun KKP dengan format berikut. Gunakan tag [align:center] untuk judul dan [align:justify] untuk narasi.

PT AGRINAS PANGAN NUSANTARA (PERSERO)
INTERNAL AUDIT (IA)

1. No. KKP              : 
2. Nama Unit Kerja      : 
3. Periode Pemeriksaan  : 
4. INTERNAL AUDITOR     : 
5. AUDITEE              : 
6. Materi Pemeriksaan   : 

**URAIAN PEMERIKSAAN**:
[align:justify]...[/align]

**CATATAN PEMERIKSA**:
[align:justify]...[/align]

**Atas Catatan Pemeriksa, Bahwa Kondisi Tersebut Belum Sesuai Dengan**:
[align:justify]...[/align]

**REKOMENDASI**:
[align:justify]...[/align]
"""

# --- FUNGSI GENERATOR ---
def generate_kkp(api_key, model_name, raw_data):
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(model_name)
        full_prompt = f"{SYSTEM_INSTRUCTION}\n\nDATA MENTAH:\n{raw_data}"
        with st.spinner(f'Menyusun dokumen menggunakan {model_name}...'):
            response = model.generate_content(full_prompt)
            return response.text
    except Exception as e:
        st.error(f"Gagal memproses AI: {e}")
        return None

# --- FUNGSI DOWNLOAD WORD (RAPI & RATA) ---
def create_docx(text):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(11)

    for line in text.split('\n'):
        if not line.strip():
            doc.add_paragraph()
            continue

        # Deteksi Header 1-6 untuk perataan titik dua
        header_match = re.match(r'^(\d\.\s[^:]+)\s*:\s*(.*)', line)
        align_match = re.match(r'^\[align:(left|center|right|justify)\](.*)\[/align\]$', line)
        
        p = doc.add_paragraph()
        
        if header_match:
            # Gunakan Tab untuk meratakan titik dua
            label = header_match.group(1)
            value = header_match.group(2)
            run_label = p.add_run(f"{label}")
            run_label.bold = True
            p.add_run(f"\t: {value}")
            # Set Tab Stop di 2 inci agar semua ':' sejajar
            p.paragraph_format.tab_stops.add_tab_stop(Inches(2))
        elif align_match:
            mode = align_match.group(1)
            content = align_match.group(2)
            if mode == 'center': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif mode == 'justify': p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
            clean_text = content.replace('**', '')
            run = p.add_run(clean_text)
            if "**" in content: run.bold = True
        else:
            clean_text = line.replace('**', '')
            run = p.add_run(clean_text)
            if "**" in line: run.bold = True
            if line.startswith('- '): p.style = 'List Bullet'

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- FUNGSI DOWNLOAD PDF ---
def create_pdf(text):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=10) # Menggunakan font standar yang stabil
    
    for line in text.split('\n'):
        if not line.strip():
            pdf.ln(5)
            continue
            
        # Bersihkan tag align untuk PDF
        align = 'L'
        if '[align:center]' in line: align = 'C'
        elif '[align:justify]' in line: align = 'J'
        
        clean_line = re.sub(r'\[.*?\]|\*\*', '', line)
        
        # Logika indentasi manual untuk header 1-6
        if re.match(r'^\d\.', clean_line):
            pdf.set_font("Arial", 'B', 10)
            parts = clean_line.split(':', 1)
            pdf.cell(50, 7, parts[0].strip(), ln=0)
            pdf.set_font("Arial", '', 10)
            if len(parts) > 1:
                pdf.cell(0, 7, f": {parts[1].strip()}", ln=1)
            else: pdf.ln(7)
        else:
            pdf.multi_cell(0, 7, clean_line, align=align)

    buf = BytesIO()
    pdf.output(buf)
    buf.seek(0)
    return buf

# --- INTERFACE ---
st.title("📄 KKP AI Generator Pro")

with st.sidebar:
    st.header("⚙️ Pengaturan")
    key = st.text_input("Gemini API Key", type="password")
    mod_opt = st.selectbox("Pilih Model", AVAILABLE_MODELS)
    model_final = st.text_input("Custom Model", "gemini-1.5-pro") if mod_opt == "Input Manual..." else mod_opt

st.subheader("1. Input Data Temuan")
raw_input = st.text_area("Masukkan catatan audit Anda di sini...", height=200, placeholder="Contoh: Unit Pemasaran, ditemukan klaim ganda 1jt...")

if st.button("🚀 Generate KKP", type="primary"):
    if not key: st.warning("Masukkan API Key!")
    else:
        res = generate_kkp(key, model_final, raw_input)
        if res:
            st.subheader("2. Preview & Download")
            st.info("Gunakan tombol di bawah untuk mengunduh hasil yang sudah dirapikan.")
            st.text_area("Preview Teks", res, height=300)
            
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("📥 Download Word (Rapi)", create_docx(res), "KKP_Final.docx")
            with c2:
                st.download_button("📥 Download PDF (Rata)", create_pdf(res), "KKP_Final.pdf")
