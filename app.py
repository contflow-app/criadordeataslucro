from __future__ import annotations

import io
import os
import re
import time
import zipfile
import tempfile
import subprocess
from dataclasses import dataclass
from datetime import datetime, date
from pathlib import Path
from typing import List, Tuple, Optional

import streamlit as st
import pdfplumber
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from docx import Document

# OCR
try:
    import pytesseract
    from pdf2image import convert_from_bytes
    OCR_OK = True
except Exception:
    OCR_OK = False

# OpenAI
try:
    from openai import OpenAI
    OPENAI_OK = True
except Exception:
    OPENAI_OK = False


# =========================
# Constantes
# =========================
DEFAULT_PRESIDENTE = "STANLEY DE SOUZA MOREIRA"
DEFAULT_ADVOGADO = "MARCO AURÉLIO POFFO"

PT_MONTHS = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
    7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}

CNPJ_RE = re.compile(r"\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b")
CPF_RE  = re.compile(r"\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b")
NIRE_RE = re.compile(r"\bNIRE\s*(?:n[º°o]?\.?)?\s*[:\-]?\s*([0-9.\-]{5,})", re.IGNORECASE)

RAZAO_RE = re.compile(
    r"([A-ZÁÂÃÉÊÍÓÔÕÚÇ0-9][A-ZÁÂÃÉÊÍÓÔÕÚÇ0-9\s\.\-&]{6,}(?:LTDA|Ltda|S\/A|SA|EIRELI|ME|EPP))"
)

ENDERECO_PATTERNS = [
    re.compile(r"localizad[ao]\s+no\s+endere[cç]o\s+(.+?)(?:\.\s|\n)", re.IGNORECASE | re.DOTALL),
    re.compile(r"tem\s+sede\s*:\s*(.+?)(?:\.\s|\n)", re.IGNORECASE | re.DOTALL),
    re.compile(r"sede\s+na\s+(.+?)(?:\.\s|\n)", re.IGNORECASE | re.DOTALL),
]


# =========================
# Modelos
# =========================
@dataclass
class Socio:
    nome: str
    cpf: str

@dataclass
class Extracao:
    razao_social: str = ""
    cnpj: str = ""
    nire: str = ""
    endereco: str = ""
    cidade_uf: str = ""
    socios: List[Socio] = None

    def __post_init__(self):
        if self.socios is None:
            self.socios = []

@dataclass
class Resultado:
    arquivo_origem: str
    status: str
    mensagem: str
    metodo: str
    ocr_usado: bool
    razao_social: str
    cnpj: str
    nire: str
    cidade_uf: str
    qtd_socios: int
    tempo_ms: int
    docx_bytes: Optional[bytes] = None
    pdf_bytes: Optional[bytes] = None


# =========================
# Utilidades
# =========================
def normalize_spaces(s: str) -> str:
    s = (s or "").replace("\u00A0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()

def city_uf_from_text(text: str) -> str:
    m = re.search(r"\b([A-ZÁÂÃÉÊÍÓÔÕÚÇ][A-Za-zÁÂÃÉÊÍÓÔÕÚÇ\s]{2,})\s*[/-]\s*([A-Z]{2})\b", text or "")
    return f"{normalize_spaces(m.group(1))}/{m.group(2)}" if m else ""

def extract_endereco(text: str) -> Tuple[str, str]:
    for pat in ENDERECO_PATTERNS:
        m = pat.search(text or "")
        if m:
            end = normalize_spaces(m.group(1))
            return end, city_uf_from_text(end)
    return "", ""

def extract_socios(text: str) -> List[Socio]:
    socios, seen = [], set()
    lines = [normalize_spaces(l) for l in (text or "").splitlines() if normalize_spaces(l)]

    for i in range(1, len(lines)):
        if "CPF" in lines[i].upper():
            m = CPF_RE.search(lines[i])
            if m:
                cpf = m.group(0)
                nome = lines[i-1]
                if cpf not in seen:
                    socios.append(Socio(nome, cpf))
                    seen.add(cpf)

    if not socios:
        for m in CPF_RE.finditer(text or ""):
            cpf = m.group(0)
            if cpf not in seen:
                socios.append(Socio("", cpf))
                seen.add(cpf)
    return socios

def regex_extract(text: str) -> Extracao:
    razao = normalize_spaces(RAZAO_RE.search(text).group(1)) if RAZAO_RE.search(text) else ""
    cnpj = CNPJ_RE.search(text).group(0) if CNPJ_RE.search(text) else ""
    nire = NIRE_RE.search(text).group(1) if NIRE_RE.search(text) else ""
    endereco, cidade_uf = extract_endereco(text)
    socios = extract_socios(text)
    return Extracao(razao, cnpj, nire, endereco, cidade_uf, socios)

def extract_text_from_pdf(bts: bytes) -> Tuple[str, bool]:
    native = ""
    try:
        with pdfplumber.open(io.BytesIO(bts)) as pdf:
            native = "\n".join(p.extract_text() or "" for p in pdf.pages)
    except Exception:
        pass
    if len(native) > 200:
        return native, False
    if OCR_OK:
        images = convert_from_bytes(bts, dpi=250)
        ocr = "\n".join(pytesseract.image_to_string(i, lang="por") for i in images)
        return (ocr if len(ocr) > len(native) else native), True
    return native, False

def extract_text_from_docx(bts: bytes) -> str:
    doc = Document(io.BytesIO(bts))
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


# =========================
# GPT
# =========================
def gpt_extract(text: str) -> Optional[Extracao]:
    client = OpenAI(api_key=st.secrets.get("OPENAI_API_KEY"))
    schema = {
        "type": "object",
        "properties": {
            "razao_social": {"type": "string"},
            "cnpj": {"type": "string"},
            "nire": {"type": "string"},
            "endereco": {"type": "string"},
            "cidade_uf": {"type": "string"},
            "socios": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {"nome": {"type": "string"}, "cpf": {"type": "string"}},
                    "required": ["nome", "cpf"]
                }
            }
        },
        "required": ["razao_social", "cnpj", "nire", "endereco", "cidade_uf", "socios"]
    }

    resp = client.responses.create(
        model="gpt-4o-mini",
        input=text[:120000],
        text={
            "format": {
                "type": "json_schema",
                "name": "extract_societario",
                "strict": True,
                "schema": schema
            }
        }
    )
    import json
    data = json.loads(resp.output_text)
    socios = [Socio(s["nome"], s["cpf"]) for s in data["socios"]]
    return Extracao(
        data["razao_social"], data["cnpj"], data["nire"],
        data["endereco"], data["cidade_uf"], socios
    )


# =========================
# DOCX
# =========================
def fill_template_docx(template_path: str, ext: Extracao, ata_date: date, ata_time: str) -> bytes:
    doc = Document(template_path)
    dia, mes, ano = str(ata_date.day), PT_MONTHS[ata_date.month], str(ata_date.year)

    for p in doc.paragraphs:
        if p.text.strip().startswith("Aos dias"):
            p.text = re.sub(
                r"Aos\s+dias.+?,",
                f"Aos dias {dia} do mês de {mes} do ano de {ano}, às {ata_time} horas,",
                p.text
            )

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# =========================
# PROCESSAMENTO
# =========================
def process_one(name, bts, template_path, ata_date, ata_time, use_gpt):
    start = time.time()
    ocr = False

    if name.lower().endswith(".pdf"):
        text, ocr = extract_text_from_pdf(bts)
    else:
        text = extract_text_from_docx(bts)

    ext = gpt_extract(text) if use_gpt else regex_extract(text)
    docx = fill_template_docx(template_path, ext, ata_date, ata_time)

    return Resultado(
        name, "OK", "Gerado", "gpt" if use_gpt else "regex", ocr,
        ext.razao_social, ext.cnpj, ext.nire, ext.cidade_uf,
        len(ext.socios), int((time.time()-start)*1000), docx, None
    )


# =========================
# UI
# =========================
st.set_page_config(layout="wide")
st.title("Criador de Atas")

with st.sidebar:
    template_path = st.text_input("Template", "templates/TEMPLATE_ATA.docx")
    ata_date = st.date_input("Data da ATA", datetime.now().date())
    use_gpt = st.checkbox("Usar GPT")

tab1, tab2 = st.tabs(["Individual", "Lote"])

with tab1:
    up = st.file_uploader("Contrato", type=["pdf", "docx"])
    if up:
        ata_time = datetime.now().strftime("%H:%M")
        r = process_one(up.name, up.read(), template_path, ata_date, ata_time, use_gpt)
        st.download_button("Baixar ATA", r.docx_bytes, file_name="ATA.docx")

with tab2:
    ups = st.file_uploader("Contratos", type=["pdf", "docx"], accept_multiple_files=True)
    if st.button("Processar lote") and ups:
        batch_time = datetime.now().strftime("%H:%M")
        for u in ups:
            r = process_one(u.name, u.read(), template_path, ata_date, batch_time, use_gpt)
            st.download_button(f"ATA {u.name}", r.docx_bytes, file_name=f"ATA_{u.name}.docx")
