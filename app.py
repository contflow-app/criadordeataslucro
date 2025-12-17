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
    t = text or ""
    m = re.search(r"\b([A-ZÁÂÃÉÊÍÓÔÕÚÇ][A-Za-zÁÂÃÉÊÍÓÔÕÚÇ\s]{2,})\s*/\s*([A-Z]{2})\b", t)
    if m:
        return f"{normalize_spaces(m.group(1))}/{m.group(2)}"
    m = re.search(r"\b([A-ZÁÂÃÉÊÍÓÔÕÚÇ][A-Za-zÁÂÃÉÊÍÓÔÕÚÇ\s]{2,})\s*-\s*([A-Z]{2})\b", t)
    if m:
        return f"{normalize_spaces(m.group(1))}/{m.group(2)}"
    return ""

def extract_endereco(text: str) -> Tuple[str, str]:
    for pat in ENDERECO_PATTERNS:
        m = pat.search(text or "")
        if m:
            end = normalize_spaces(m.group(1))
            return end, city_uf_from_text(end)
    return "", ""

def extract_socios(text: str) -> List[Socio]:
    """
    1) assinatura: linha nome + linha CPF
    2) inline CPF
    3) fallback CPFs
    """
    t = text or ""
    lines = [normalize_spaces(l) for l in t.splitlines() if normalize_spaces(l)]
    socios: List[Socio] = []
    seen = set()

    for i in range(1, len(lines)):
        if "CPF" in lines[i].upper():
            m = CPF_RE.search(lines[i])
            if not m:
                continue
            cpf = m.group(0)
            nome = lines[i - 1]
            if cpf not in seen:
                socios.append(Socio(nome=nome, cpf=cpf))
                seen.add(cpf)

    if not socios:
        for m in re.finditer(r"([A-ZÁÂÃÉÊÍÓÔÕÚÇ][A-ZÁÂÃÉÊÍÓÔÕÚÇ\s]{4,120}).{0,120}?\bCPF\b.{0,40}?("+CPF_RE.pattern+r")", t):
            nome = normalize_spaces(m.group(1))
            cpf = m.group(2)
            if cpf not in seen:
                socios.append(Socio(nome=nome, cpf=cpf))
                seen.add(cpf)

    if not socios:
        for m in CPF_RE.finditer(t):
            cpf = m.group(0)
            if cpf not in seen:
                socios.append(Socio(nome="", cpf=cpf))
                seen.add(cpf)

    return socios

def regex_extract(text: str) -> Extracao:
    razao = normalize_spaces(RAZAO_RE.search(text).group(1)) if RAZAO_RE.search(text) else ""
    cnpj = CNPJ_RE.search(text).group(0) if CNPJ_RE.search(text) else ""
    nire = NIRE_RE.search(text).group(1) if NIRE_RE.search(text) else ""
    endereco, cidade_uf = extract_endereco(text)
    socios = extract_socios(text)
    return Extracao(razao_social=razao, cnpj=cnpj, nire=nire, endereco=endereco, cidade_uf=cidade_uf, socios=socios)

def extract_text_from_pdf(bts: bytes, min_chars: int = 250) -> Tuple[str, bool]:
    native = ""
    try:
        parts = []
        with pdfplumber.open(io.BytesIO(bts)) as pdf:
            for page in pdf.pages:
                t = (page.extract_text() or "").strip()
                if t:
                    parts.append(t)
        native = "\n".join(parts).strip()
    except Exception:
        native = ""

    if len(native) >= min_chars:
        return native, False

    if not OCR_OK:
        return native, False

    try:
        images = convert_from_bytes(bts, dpi=250)
        ocr_parts = []
        for img in images:
            txt = pytesseract.image_to_string(img, lang="por", config="--oem 3 --psm 6")
            txt = (txt or "").strip()
            if txt:
                ocr_parts.append(txt)
        ocr_text = "\n".join(ocr_parts).strip()
        if len(ocr_text) > len(native):
            return ocr_text, True
        return native, False
    except Exception:
        return native, False

def extract_text_from_docx(bts: bytes) -> str:
    doc = Document(io.BytesIO(bts))
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


# =========================
# GPT (Structured Outputs)
# =========================
def gpt_extract(text: str) -> Optional[Extracao]:
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY ausente em Secrets.")
    if not OPENAI_OK:
        raise RuntimeError("Biblioteca openai não instalada (requirements.txt).")

    client = OpenAI(api_key=api_key)

    schema = {
        "type": "object",
        "additionalProperties": False,
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
                    "additionalProperties": False,
                    "properties": {
                        "nome": {"type": "string"},
                        "cpf": {"type": "string"},
                    },
                    "required": ["nome", "cpf"],
                }
            }
        },
        "required": ["razao_social", "cnpj", "nire", "endereco", "cidade_uf", "socios"]
    }

    prompt = f"""
Extraia dados societários do texto abaixo e retorne APENAS JSON no schema.
Regras:
- Não invente dados.
- Liste TODOS os sócios com NOME e CPF quando disponíveis. Se CPF não aparecer, cpf="".
- Não confunda NIRE com CPF.
- cidade_uf apenas se explícito.
- Se algo não estiver no texto, retorne "".

TEXTO:
\"\"\"{text[:120000]}\"\"\"
"""

    resp = client.responses.create(
        model="gpt-4o-mini",
        input=[{"role": "user", "content": prompt}],
        text={
            "format": {
                "type": "json_schema",
                "name": "extract_societario",
                "strict": True,
                "schema": schema,
            }
        }
    )

    raw = getattr(resp, "output_text", "") or ""
    import json
    data = json.loads(raw)

    socios = [Socio(nome=s.get("nome", ""), cpf=s.get("cpf", "")) for s in (data.get("socios") or [])]
    return Extracao(
        razao_social=data.get("razao_social", ""),
        cnpj=data.get("cnpj", ""),
        nire=data.get("nire", ""),
        endereco=data.get("endereco", ""),
        cidade_uf=data.get("cidade_uf", ""),
        socios=socios,
    )


# =========================
# DOCX helpers (placeholders)
# =========================
def replace_placeholders_in_paragraph(paragraph, mapping: dict) -> None:
    """
    Substitui placeholders em runs, preservando formatação, mas com uma limitação:
    placeholders não devem estar "quebrados" em múltiplos runs.
    (No template melhorado, isso é controlável.)
    """
    for run in paragraph.runs:
        text = run.text
        if not text:
            continue
        changed = False
        for k, v in mapping.items():
            token = "{{" + k + "}}"
            if token in text:
                text = text.replace(token, v)
                changed = True
        if changed:
            run.text = text

def set_multiline_placeholder(paragraph, token_name: str, lines: List[str]) -> None:
    """
    Substitui um parágrafo que contém {{TOKEN}} por múltiplas linhas
    (com quebras de linha dentro do mesmo parágrafo).
    """
    token = "{{" + token_name + "}}"
    if token not in paragraph.text:
        return

    # limpa runs e recria com breaks
    paragraph.clear()
    if not lines:
        paragraph.add_run("________________________")
        return

    for i, line in enumerate(lines):
        r = paragraph.add_run(line)
        if i < len(lines) - 1:
            r.add_break()

def fill_template_docx(template_path: str, ext: Extracao, ata_date: date, ata_time: str) -> bytes:
    doc = Document(template_path)

    dia = str(ata_date.day)
    mes = PT_MONTHS[ata_date.month]
    ano = str(ata_date.year)

    # monta listas
    socios_presenca_lines = []
    for idx, s in enumerate(ext.socios or [], start=1):
        nome = (s.nome or "").strip() or "________________________"
        socios_presenca_lines.append(f"{idx}. {nome}")

    socios_assinaturas_lines = []
    for s in (ext.socios or []):
        nome = (s.nome or "").strip() or "________________________"
        cpf = (s.cpf or "").strip()
        socios_assinaturas_lines.append("_______________________________________")
        socios_assinaturas_lines.append(nome)
        if cpf:
            socios_assinaturas_lines.append(f"CPF nº {cpf}")
        socios_assinaturas_lines.append("")  # linha em branco entre sócios

    mapping = {
        "RAZAO_SOCIAL": ext.razao_social or "________________________",
        "CNPJ": ext.cnpj or "____________________",
        "NIRE": ext.nire or "____________",
        "ENDERECO": ext.endereco or "__________________________________________",
        "CIDADE_UF": ext.cidade_uf or "________________",
        "DIA": dia,
        "MES": mes,
        "ANO": ano,
        "HORA": ata_time,
        "PRESIDENTE": DEFAULT_PRESIDENTE,
    }

    for p in doc.paragraphs:
        # placeholders simples
        replace_placeholders_in_paragraph(p, mapping)

        # placeholders multi-linha
        set_multiline_placeholder(p, "SOCIOS_PRESENCA", socios_presenca_lines)
        set_multiline_placeholder(p, "SOCIOS_ASSINATURAS", socios_assinaturas_lines)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# =========================
# PDF export (opcional)
# =========================
def docx_to_pdf_via_libreoffice(docx_bytes: bytes) -> Optional[bytes]:
    with tempfile.TemporaryDirectory() as tmp:
        docx_path = os.path.join(tmp, "ata.docx")
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)

        cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, docx_path]
        try:
            subprocess.run(cmd, check=True, capture_output=True)
            pdf_path = os.path.join(tmp, "ata.pdf")
            if not os.path.exists(pdf_path):
                return None
            with open(pdf_path, "rb") as f:
                return f.read()
        except Exception:
            return None


# =========================
# Excel report
# =========================
def make_report_xlsx(rows: List[Resultado]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório"
    headers = [
        "arquivo_origem","status","mensagem","ocr_usado","metodo",
        "razao_social","cnpj","nire","cidade_uf","qtd_socios",
        "docx_gerado","pdf_gerado","tempo_ms"
    ]
    ws.append(headers)
    for r in rows:
        ws.append([
            r.arquivo_origem, r.status, r.mensagem,
            "sim" if r.ocr_usado else "não",
            r.metodo,
            r.razao_social, r.cnpj, r.nire, r.cidade_uf, r.qtd_socios,
            "sim" if r.docx_bytes else "não",
            "sim" if r.pdf_bytes else "não",
            r.tempo_ms
        ])
    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 22
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


def load_batch_files(uploaded_files, uploaded_zip) -> List[Tuple[str, bytes]]:
    files: List[Tuple[str, bytes]] = []
    if uploaded_files:
        for f in uploaded_files:
            files.append((f.name, f.read()))
    if uploaded_zip:
        zbytes = uploaded_zip.read()
        with zipfile.ZipFile(io.BytesIO(zbytes)) as z:
            for name in z.namelist():
                if name.endswith("/"):
                    continue
                if name.lower().endswith((".pdf", ".docx")):
                    files.append((Path(name).name, z.read(name)))
    return files


# =========================
# Processamento
# =========================
def process_one(
    name: str,
    bts: bytes,
    template_path: str,
    ata_date: date,
    ata_time: str,
    try_pdf: bool,
    use_gpt: bool
) -> Resultado:
    t0 = time.time()
    ocr_used = False
    metodo = "regex"
    try:
        if name.lower().endswith(".pdf"):
            text, ocr_used = extract_text_from_pdf(bts)
        elif name.lower().endswith(".docx"):
            text = extract_text_from_docx(bts)
        else:
            return Resultado(name, "ERRO", "Formato não suportado (use PDF/DOCX).", "n/a", False, "", "", "", "", 0, int((time.time()-t0)*1000))

        if not text.strip():
            return Resultado(name, "ERRO", "Texto vazio (OCR falhou/indisponível).", "n/a", ocr_used, "", "", "", "", 0, int((time.time()-t0)*1000))

        if use_gpt:
            ext = gpt_extract(text)
            metodo = "gpt"
        else:
            ext = regex_extract(text)

        issues = []
        if not ext.razao_social:
            issues.append("Razão social ausente")
        if not ext.cnpj:
            issues.append("CNPJ ausente")
        if not (ext.socios and any((s.nome or "").strip() for s in ext.socios)):
            issues.append("Nomes de sócios ausentes")

        status = "OK" if not issues else "PENDENTE"
        msg = "Gerado" if status == "OK" else "; ".join(issues)
        if ocr_used:
            msg += " (OCR aplicado)"
        if use_gpt:
            msg += " (GPT)"

        docx_bytes = fill_template_docx(template_path, ext, ata_date, ata_time)
        pdf_bytes = docx_to_pdf_via_libreoffice(docx_bytes) if try_pdf else None

        return Resultado(
            arquivo_origem=name,
            status=status,
            mensagem=msg,
            metodo=metodo,
            ocr_usado=ocr_used,
            razao_social=ext.razao_social,
            cnpj=ext.cnpj,
            nire=ext.nire,
            cidade_uf=ext.cidade_uf,
            qtd_socios=len(ext.socios or []),
            tempo_ms=int((time.time()-t0)*1000),
            docx_bytes=docx_bytes,
            pdf_bytes=pdf_bytes
        )
    except Exception as e:
        return Resultado(name, "ERRO", f"{type(e).__name__}: {e}", metodo, ocr_used, "", "", "", "", 0, int((time.time()-t0)*1000))


# =========================
# UI
# =========================
st.set_page_config(page_title="IA para geração de Atas - Lei 15270/25", layout="wide")
st.markdown(
    """
    <div style="text-align:center; margin-bottom:20px;">
        <h1>Criador de Atas</h1>
        <div style="color:#777; font-size:0.9em;">
            Desenvolvido por <strong>Raul Martins</strong>
        </div>
    </div>
    <hr/>
    """,
    unsafe_allow_html=True
)

with st.sidebar:
    st.subheader("Template")
    template_path = st.text_input("Caminho do template", value="templates/TEMPLATE_ATA_MELHORADO.docx")
    try_pdf = st.checkbox("Gerar PDF (requer LibreOffice)", value=False)

    st.subheader("Data da ATA")
    ata_date = st.date_input("Data (padrão = hoje)", value=datetime.now().date())

    st.subheader("Extração")
    use_gpt = st.checkbox("Usar GPT para melhorar extração", value=False)
    if use_gpt:
        if not OPENAI_OK:
            st.error("Biblioteca openai não instalada (requirements.txt).")
        elif not st.secrets.get("OPENAI_API_KEY", ""):
            st.warning("Falta OPENAI_API_KEY em Secrets do Streamlit Cloud.")

tab1, tab2 = st.tabs(["Individual", "Lote"])

with tab1:
    st.subheader("Individual")
    up = st.file_uploader("Suba 1 contrato (PDF ou DOCX)", type=["pdf", "docx"])
    if up:
        single_time = datetime.now().strftime("%H:%M")
        r = process_one(up.name, up.read(), template_path, ata_date, single_time, try_pdf, use_gpt)

        st.write(f"**Status:** {r.status} | **Método:** {r.metodo}")
        st.write(f"**Mensagem:** {r.mensagem}")
        st.write(f"**Sócios encontrados:** {r.qtd_socios}")
        st.caption(f"Carimbo: {ata_date.strftime('%d/%m/%Y')} {single_time}")

        base = Path(up.name).stem
        if r.docx_bytes:
            st.download_button(
                "Baixar ATA (DOCX)",
                data=r.docx_bytes,
                file_name=f"ATA_{base}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        if r.pdf_bytes:
            st.download_button(
                "Baixar ATA (PDF)",
                data=r.pdf_bytes,
                file_name=f"ATA_{base}.pdf",
                mime="application/pdf"
            )

with tab2:
    st.subheader("Lote")
    uploaded_files = st.file_uploader("Vários arquivos (PDF/DOCX)", type=["pdf", "docx"], accept_multiple_files=True)
    uploaded_zip = st.file_uploader("Ou ZIP com PDFs/DOCXs", type=["zip"])
    batch = load_batch_files(uploaded_files, uploaded_zip)
    st.write(f"Arquivos carregados: **{len(batch)}**")

    continue_on_error = st.checkbox("Continuar mesmo se houver erro", value=True)
    include_excel_inside_zip = st.checkbox("Incluir Excel dentro do ZIP", value=True)

    if st.button("Processar lote", disabled=(len(batch) == 0)):
        # trava carimbo no início do lote
        batch_date = ata_date
        batch_time = datetime.now().strftime("%H:%M")
        st.info(f"Carimbo do lote: {batch_date.strftime('%d/%m/%Y')} {batch_time}")

        rows: List[Resultado] = []
        prog = st.progress(0.0)
        box = st.empty()

        for i, (name, bts) in enumerate(batch, start=1):
            box.write(f"Processando {i}/{len(batch)}: **{name}**")
            r = process_one(name, bts, template_path, batch_date, batch_time, try_pdf, use_gpt)
            rows.append(r)
            prog.progress(i / len(batch))
            if (not continue_on_error) and r.status == "ERRO":
                st.error(f"Parado no erro: {name} — {r.mensagem}")
                break

        excel_bytes = make_report_xlsx(rows)

        zip_out = io.BytesIO()
        with zipfile.ZipFile(zip_out, "w", zipfile.ZIP_DEFLATED) as z:
            for r in rows:
                base = Path(r.arquivo_origem).stem
                if r.docx_bytes:
                    z.writestr(f"ATAs/{base}.docx", r.docx_bytes)
                if r.pdf_bytes:
                    z.writestr(f"ATAs/{base}.pdf", r.pdf_bytes)
            if include_excel_inside_zip:
                z.writestr("relatorio_processamento.xlsx", excel_bytes)

        zip_out.seek(0)

        ok = sum(1 for r in rows if r.status == "OK")
        pend = sum(1 for r in rows if r.status == "PENDENTE")
        err = sum(1 for r in rows if r.status == "ERRO")
        st.success(f"Concluído. OK: {ok} | PENDENTE: {pend} | ERRO: {err}")

        st.download_button(
            "Baixar Relatório (Excel)",
            data=excel_bytes,
            file_name="relatorio_processamento.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            "Baixar ZIP (ATAs + Excel opcional)",
            data=zip_out.getvalue(),
            file_name="atas_geradas.zip",
            mime="application/zip",
        )

        st.dataframe([{
            "arquivo": r.arquivo_origem,
            "status": r.status,
            "metodo": r.metodo,
            "ocr": "sim" if r.ocr_usado else "não",
            "socios": r.qtd_socios,
            "mensagem": r.mensagem
        } for r in rows])
