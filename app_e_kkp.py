import streamlit as st
import os
import re
import google.generativeai as genai
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
from io import BytesIO

# --- KONFIGURASI HALAMAN WEB ---
st.set_page_config(page_title="KKP Generator AI", layout="wide")

# --- DAFTAR MODEL YANG TERSEDIA ---
# Anda bisa menambahkan atau mengurangi daftar ini sesuai kebutuhan
AVAILABLE_MODELS = [
    "gemini-1.5-flash",          # Model standar cepat & stabil
    "gemini-1.5-pro",            # Model standar high-intelligence
    "gemini-2.0-flash-exp",      # (Opsional: Versi eksperimental terbaru saat ini)
    "gemini-2.5-flash",          # Sesuai request
    "gemini-3-flash-preview",    # Sesuai request
    "gemini-3-pro-preview",      # Sesuai request
    "Input Manual..."            # Opsi jika ingin mengetik nama model lain
]

# --- INSTRUKSI SISTEM ---
SYSTEM_INSTRUCTION = """
Anda adalah Auditor Senior di Internal Audit (IA) PT AGRINAS PANGAN NUSANTARA (PERSERO).
Tugas Anda adalah menyusun Kertas Kerja Pemeriksaan (KKP) berdasarkan data temuan mentah yang diberikan.
Anda WAJIB mengikuti format formulir standar dan prosedur berikut:
FORMAT FORMULIR (OUTPUT YANG DIINGINKAN):
---------------------------------------------------------
[align:center]PT AGRINAS PANGAN NUSANTARA (PERSERO)[/align]
[align:center]INTERNAL AUDIT (IA)[/align]

1. No. KKP              : [Isi atau strip "-" jika kosong]
2. Nama Unit Kerja      : [Isi Nama Unit/Divisi]
3. Periode Pemeriksaan  : [Isi Tanggal]
4. INTERNAL AUDITOR     : [Sebutkan nama-nama tim]
5. AUDITEE              : [Sebutkan nama auditee/pihak yang diperiksa]
6. Materi Pemeriksaan   : [Judul Audit]

**URAIAN PEMERIKSAAN**:
[align:justify][Tulis paragraf singkat tentang apa yang diperiksa][/align]

**CATATAN PEMERIKSA**:
[align:justify][Buat poin-poin (bullet points). Jelaskan temuan masalah secara detail. Jika ada angka uang (Rp), tuliskan dengan jelas][/align]

**Atas Catatan Pemeriksa, Bahwa Kondisi Tersebut Belum Sesuai Dengan**:
[align:justify][Sebutkan peraturan-peraturan yang dilanggar (misalnya, PKB Bab XI Pasal 59). Jika tidak disebutkan dalam data mentah, kosongkan atau isi dengan "..."][/align]

**Kondisi Tersebut Dapat Mengakibatkan**:
[align:justify][Secara otomatis diisi oleh AI mengenai akibat atau dampak dari permasalahan tersebut.][/align]

**Kondisi Tersebut Disebabkan Oleh**:
[align:justify][Secara otomatis diisi oleh AI mengenai penyebab atau akar masalah dari permasalahan tersebut.][/align]

**Analisis Governance, Risk dan Compliance (GRC)**:
[align:justify][Secara otomatis diisi oleh AI mengenai analisis tata kelola, risiko, dan kepatuhan terkait temuan.][/align]

**REKOMENDASI**:
[align:justify][Tuliskan saran perbaikan. Jika user menyebutkan aturan/pasal, sertakan di sini][/align]
---------------------------------------------------------
ATURAN PENGISIAN:
1. Gunakan Bahasa Indonesia formal dan istilah audit yang tepat.
2. Gunakan tag [align:justify]...[/align] untuk paragraf narasi.
3. Jangan mengarang data yang tidak ada di sumber input.
"""

# --- FUNGSI GENERATOR AI (DIPERBARUI DENGAN PILIHAN MODEL) ---
def generate_kkp_from_gemini(api_key, model_name, raw_data):
    try:
        genai.configure(api_key=api_key)
        
        # Inisialisasi model sesuai pilihan user
        model = genai.GenerativeModel(model_name)
        
        prompt = f"""
        Buatkan dokumen KKP lengkap berdasarkan data mentah berikut ini.
        Pastikan tata letak output meniru formulir resmi perusahaan.

        DATA MENTAH:
        {raw_data}
        """
        full_prompt = f"{SYSTEM_INSTRUCTION}\n{prompt}"
        
        with st.spinner(f'Sedang memproses menggunakan model: {model_name}...'):
            # Menambahkan config untuk memastikan output lebih konsisten
            response = model.generate_content(full_prompt)
            return response.text
    except Exception as e:
        # Menangani error jika model belum tersedia di API Key user
        if "404" in str(e) or "not found" in str(e).lower():
            st.error(f"Error: Model '{model_name}' tidak ditemukan atau belum tersedia untuk API Key Anda. Coba gunakan 'gemini-1.5-flash'.")
        else:
            st.error(f"Error pada Gemini API: {e}")
        return None

# --- FUNGSI PEMBUAT DOCX ---
def create_docx_in_memory(formatted_text):
    doc = Document()
    lines = formatted_text.split('\n')
    
    for line in lines:
        if line.strip() == '':
            doc.add_paragraph()
            continue
            
        align_match = re.match(r'^\[align:(left|center|right|justify)\](.*)\[/align\]$', line)
        content = line
        alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        if align_match:
            align_type = align_match.group(1)
            content = align_match.group(2)
            if align_type == 'center': alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif align_type == 'right': alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif align_type == 'left': alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        p = doc.add_paragraph()
        p.alignment = alignment
        
        if content.startswith('- '):
            p.style = 'List Bullet'
            content = content[2:]
            
        clean_content = content.replace('**', '').replace('*', '') 
        run = p.add_run(clean_content)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        if "**" in line:
             run.bold = True

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- FUNGSI PEMBUAT PDF ---
class PDF(FPDF):
    def header(self): pass
    def footer(self): pass

def create_pdf_in_memory(formatted_text):
    pdf = PDF('P', 'mm', 'A4')
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.set_font("Times", size=11)
    
    line_height = 6
    
    for line in formatted_text.split('\n'):
        if line.strip() == '':
            pdf.ln(line_height * 0.5)
            continue

        align_match = re.match(r'^\[align:(left|center|right|justify)\](.*)\[/align\]$', line)
        content = line
        alignment = 'J'
        
        if align_match:
            align_type = align_match.group(1)
            content = align_match.group(2)
            if align_type == 'center': alignment = 'C'
            
        clean_text = re.sub(r'(\*\*|\*|__|\[.*?\])', '', content).strip()
        
        if content.startswith('- '):
            pdf.set_x(20)
            pdf.multi_cell(0, line_height, f"• {clean_text[2:]}", align='L')
        else:
            if "**" in content:
                pdf.set_font("Times", 'B', 11)
                pdf.multi_cell(0, line_height, clean_text, align=alignment)
                pdf.set_font("Times", '', 11)
            else:
                pdf.multi_cell(0, line_height, clean_text, align=alignment)

    try:
        pdf_output = pdf.output(dest='S').encode('latin-1', 'replace')
        return BytesIO(pdf_output)
    except Exception as e:
        st.error(f"Gagal membuat PDF: {e}")
        return None

# --- UI UTAMA (STREAMLIT) ---

st.title("📄 KKP Generator AI (Audit Toolkit)")
st.markdown("Aplikasi untuk membuat draft **Kertas Kerja Pemeriksaan** secara otomatis.")

# Sidebar untuk konfigurasi
with st.sidebar:
    st.header("Konfigurasi")
    api_key_input = st.text_input("Masukkan Google Gemini API Key", type="password")
    
    st.divider()
    
    # --- FITUR BARU: PILIHAN MODEL ---
    st.subheader("Pilih Model AI")
    selected_option = st.selectbox(
        "Versi Model Gemini:",
        AVAILABLE_MODELS,
        index=0 # Default ke opsi pertama (gemini-1.5-flash)
    )
    
    # Logika jika user memilih "Input Manual"
    if selected_option == "Input Manual...":
        final_model_name = st.text_input("Ketik kode model manual:", "gemini-1.5-pro-latest")
    else:
        final_model_name = selected_option
        
    st.info(f"Model aktif: **{final_model_name}**")
    st.caption("Pastikan API Key Anda mendukung model yang dipilih.")

# Area Input Data
st.subheader("1. Masukkan Data Temuan Audit")

raw_data_input = st.text_area(
    "Catatan Lapangan / Data Mentah",
    value="",
    height=200,
    placeholder="Contoh Input:\n- Unit: Divisi Pemasaran\n- Periode: Juni 2024\n- Temuan: Ada selisih kas sebesar Rp 1.000.000...\n(Silakan ketik temuan Anda di sini)"
)

# Tombol Generate
if st.button("🚀 Buat Draft KKP", type="primary"):
    if not api_key_input:
        st.warning("Mohon masukkan API Key terlebih dahulu di sidebar sebelah kiri.")
    elif not raw_data_input:
        st.warning("Data temuan tidak boleh kosong.")
    else:
        # 1. Proses AI (Mengirim parameter final_model_name)
        result_text = generate_kkp_from_gemini(api_key_input, final_model_name, raw_data_input)
        
        if result_text:
            st.subheader("2. Hasil Preview AI")
            st.text_area("Output Teks", value=result_text, height=400)
            
            # 2. Siapkan File Download
            col1, col2 = st.columns(2)
            
            # Buat DOCX
            docx_file = create_docx_in_memory(result_text)
            with col1:
                st.download_button(
                    label="📥 Download Word (.docx)",
                    data=docx_file,
                    file_name=f"KKP_Draft_{final_model_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            
            # Buat PDF
            pdf_file = create_pdf_in_memory(result_text)
            if pdf_file:
                with col2:
                    st.download_button(
                        label="📥 Download PDF (.pdf)",
                        data=pdf_file,
                        file_name=f"KKP_Draft_{final_model_name}.pdf",
                        mime="application/pdf"
                    )