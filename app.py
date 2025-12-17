
import io
import os
import re
import time
import zipfile
import tempfile
import subprocess
from pathlib import Path
from datetime import datetime

import streamlit as st
import pdfplumber
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from docx import Document

# OCR fallback
try:
    import pytesseract
    from pdf2image import convert_from_bytes
    OCR_OK = True
except Exception:
    OCR_OK = False

# OpenAI (optional, for better extraction)
try:
    from openai import OpenAI
    OPENAI_OK = True
except Exception:
    OPENAI_OK = False

NOW = datetime.now()

CNPJ_RE = re.compile(r"\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b")
CPF_RE  = re.compile(r"\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b")
NIRE_RE = re.compile(r"\bNIRE\s*(?:n[º°o]?\.?)?\s*[:\-]?\s*([0-9.\-]{5,})", re.IGNORECASE)

DEFAULT_PRESIDENTE = "STANLEY DE SOUZA MOREIRA"
DEFAULT_ADVOGADO   = "MARCO AURÉLIO POFFO"

PT_MONTHS = {
    1:"janeiro",2:"fevereiro",3:"março",4:"abril",5:"maio",6:"junho",
    7:"julho",8:"agosto",9:"setembro",10:"outubro",11:"novembro",12:"dezembro"
}
def today_defaults():
    d = datetime.now()
    return str(d.day), PT_MONTHS[d.month], str(d.year)


def normalize_spaces(s: str) -> str:
    s = (s or "").replace("\u00A0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()

def extract_text_from_pdf(file_bytes: bytes) -> tuple[str, bool]:
    """(text, ocr_used)"""
    native = ""
    try:
        parts = []
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                t = (page.extract_text() or "").strip()
                if t:
                    parts.append(t)
        native = "\n".join(parts).strip()
    except Exception:
        native = ""

    if len(native) >= 200:
        return native, False

    if not OCR_OK:
        return native, False

    try:
        images = convert_from_bytes(file_bytes, dpi=250)
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

def extract_text_from_docx(file_bytes: bytes) -> str:
    doc = Document(io.BytesIO(file_bytes))
    parts = []
    for p in doc.paragraphs:
        if p.text and p.text.strip():
            parts.append(p.text)
    return "\n".join(parts)

def regex_extract(text: str) -> dict:
    cnpj = CNPJ_RE.search(text or "")
    nire = NIRE_RE.search(text or "")
    # naive company name: first long uppercase line with LTDA/SA
    razao = ""
    m = re.search(r"([A-ZÁÂÃÉÊÍÓÔÕÚÇ0-9][A-ZÁÂÃÉÊÍÓÔÕÚÇ0-9\s\.\-&]+(?:LTDA|S\/A|SA|EIRELI|ME|EPP))", text or "")
    if m:
        razao = normalize_spaces(m.group(1))[:180]

    # address
    endereco = ""
    m = re.search(r"localizad[ao]\s+no\s+endere[cç]o\s+(.+?)(?:\.\s|\n)", text or "", flags=re.IGNORECASE|re.DOTALL)
    if not m:
        m = re.search(r"tem\s+sede\s*:\s*(.+?)(?:\.\s|\n)", text or "", flags=re.IGNORECASE|re.DOTALL)
    if m:
        endereco = normalize_spaces(m.group(1))

    # city/uf from address (conservative)
    cidade_uf = ""
    m = re.search(r"\b([A-ZÁÂÃÉÊÍÓÔÕÚÇ][A-Za-zÁÂÃÉÊÍÓÔÕÚÇ\s]{2,})\s*/\s*([A-Z]{2})\b", endereco)
    if m:
        cidade_uf = f"{normalize_spaces(m.group(1))}/{m.group(2)}"

    # partners: signature pair + inline CPF
    socios = []
    seen = set()
    lines = [normalize_spaces(l) for l in (text or "").splitlines() if normalize_spaces(l)]
    for i in range(1, len(lines)):
        if re.search(r"\bCPF\b", lines[i], flags=re.IGNORECASE):
            mcpf = CPF_RE.search(lines[i])
            if mcpf:
                cpf = mcpf.group(0)
                name = lines[i-1]
                if cpf not in seen and name and len(name) >= 5:
                    socios.append({"nome": name, "cpf": cpf})
                    seen.add(cpf)

    if not socios:
        for m in re.finditer(r"([A-ZÁÂÃÉÊÍÓÔÕÚÇ][A-ZÁÂÃÉÊÍÓÔÕÚÇ\s]{4,120}).{0,80}?\bCPF\b.{0,20}?("+CPF_RE.pattern+r")", text or ""):
            nome = normalize_spaces(m.group(1))
            cpf = m.group(2)
            if cpf not in seen:
                socios.append({"nome": nome, "cpf": cpf})
                seen.add(cpf)

    return {
        "razao_social": razao,
        "cnpj": cnpj.group(0) if cnpj else "",
        "nire": nire.group(1) if nire else "",
        "endereco": endereco,
        "cidade_uf": cidade_uf,
        "socios": socios
    }

def gpt_extract(text: str) -> tuple[dict, str]:
    """
    Uses OpenAI structured outputs (Responses API) to extract fields.
    Returns (data, raw_json_text)
    """
    if not OPENAI_OK:
        return {}, "openai lib not installed"
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        return {}, "missing OPENAI_API_KEY in secrets"
    client = OpenAI(api_key=api_key)

    schema = {
      "name": "extract_societario",
      "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
          "razao_social": {"type":"string"},
          "cnpj": {"type":"string"},
          "nire": {"type":"string"},
          "endereco": {"type":"string"},
          "cidade_uf": {"type":"string"},
          "socios": {
            "type":"array",
            "items":{
              "type":"object",
              "additionalProperties": False,
              "properties":{
                "nome":{"type":"string"},
                "cpf":{"type":"string"}
              },
              "required":["nome","cpf"]
            }
          }
        },
        "required":["razao_social","cnpj","nire","endereco","cidade_uf","socios"]
      },
      "strict": True
    }

    prompt = f"""
Extraia dados societários do texto abaixo e retorne JSON estritamente no schema.
Regras:
- Não invente dados. Se não houver, retorne "".
- Para sócios, tente listar TODOS com nome e CPF. Se CPF não aparecer, deixe cpf="".
- Não confunda NIRE com CPF.
- cidade_uf somente se explícito.

TEXTO:
\"\"\"{text[:120000]}\"\"\"
"""

    resp = client.responses.create(
        model="gpt-4o-mini",
        input=[{"role":"user","content":prompt}],
        text={"format":{"type":"json_schema","json_schema":schema}}
    )
    # The SDK returns output text in resp.output_text
    raw = getattr(resp, "output_text", "") or ""
    import json
    try:
        data = json.loads(raw)
    except Exception:
        data = {}
    return data, raw

def _remove_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None

def fill_template_docx(template_path: str, ctx: dict) -> bytes:
    """
    Uses the provided template (already formatted) and fills:
    - razao social, cnpj, nire, endereco
    - presença list
    - presidente/secretário defaults
    - signatures rebuilt from socios
    """
    doc = Document(template_path)

    razao = ctx.get("razao_social","")
    cnpj = ctx.get("cnpj","")
    nire = ctx.get("nire","")
    endereco = ctx.get("endereco","")
    cidade_uf = ctx.get("cidade_uf","")

    presidente = ctx.get("presidente") or DEFAULT_PRESIDENTE
    secretario = ctx.get("secretario") or DEFAULT_ADVOGADO  # user asked "pelo Advogado" - use as secretary line if needed
    socios = ctx.get("socios") or []

    # 1) Replace header paragraph pieces conservatively
    for p in doc.paragraphs:
        t = p.text or ""
        if "Sociedade empresária limitada" in t and "CNPJ/MF" in t:
            # replace company name after dash up to comma
            t2 = re.sub(r"–\s*.+?,\s*inscrita", f"– {razao or '________________________'}, inscrita", t)
            t2 = re.sub(r"CNPJ/MF\s*n[º°]\.\s*[^–]+–", f"CNPJ/MF nº. {cnpj or '____________________'} –", t2)
            if "NIRE" in t2:
                t2 = re.sub(r"NIRE:\s*[^ ]+", f"NIRE: {nire or '____________'}", t2)
            # address line
            t2 = re.sub(r"localizada no endereço\s+_+\.", f"localizada no endereço {endereco or '__________________________________________'}.", t2)
            p.text = t2

        # opening paragraph with "na sede da ..."
        dia = ctx.get("dia") or str(NOW.day)
        mes = ctx.get("mes") or "__________"
        ano = ctx.get("ano") or str(NOW.year)

        if t.startswith("Aos dias") and "na sede da" in t:
            t2 = re.sub(r"Aos\s+dias\s+.+?\s+do\s+mês\s+de\s+.+?\s+do\s+ano\s+de\s+.+?,", f"Aos dias {dia} do mês de {mes} do ano de {ano},", t)
            t = t2
            # replace company mention and address blanks
            t2 = re.sub(r"na sede da\s+.+?,\s+localizada", f"na sede da {razao or '________________________'}, localizada", t)
            t2 = re.sub(r"localizada no endereço\s*\.?", f"localizada no endereço {endereco or '__________________________________________'}.", t2)
            p.text = t2

        # mesa composition
        if t.startswith("DA COMPOSIÇÃO DA MESA"):
            # keep template wording but fill names
            t2 = re.sub(r"presidida por\s+_+", f"presidida por {presidente}", t)
            t2 = re.sub(r"SECRETÁRIO\s+_+", f"SECRETÁRIO {secretario}", t2)
            p.text = t2

        # date line with city placeholder
        if re.match(r"^_+,", t.strip()):
            # try to fill date line if template has placeholders
            if mes and ano:
                t = re.sub(r"de\s+_+\s+de\s+_+\.?$", f"de {mes} de {ano}.", t)
                t = re.sub(r"\b\d{4}\b", ano, t)

            if cidade_uf:
                p.text = re.sub(r"^_+", cidade_uf.split("/")[0], t)

    # 2) Fill presence list between "DA PRESENÇA" and "Os quais"
    pres_idx = None
    osquais_idx = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith("DA PRESENÇA"):
            pres_idx = i
        if pres_idx is not None and p.text.strip().startswith("Os quais"):
            osquais_idx = i
            break

    if pres_idx is not None and osquais_idx is not None:
        # remove existing numbered placeholders in between (lines that start with digit.)
        to_remove = []
        for j in range(pres_idx+1, osquais_idx):
            if re.match(r"^\d+\.\s*", doc.paragraphs[j].text.strip()):
                to_remove.append(doc.paragraphs[j])
        for par in to_remove:
            _remove_paragraph(par)

        # insert fresh list right after pres_idx
        insert_at = pres_idx + 1
        # python-docx cannot truly insert; we add then move by xml
        # We'll append at end of document then move.
        from docx.oxml import OxmlElement
        from docx.text.paragraph import Paragraph

        def insert_paragraph_after(paragraph, text):
            new_p = OxmlElement("w:p")
            paragraph._p.addnext(new_p)
            new_para = Paragraph(new_p, paragraph._parent)
            new_para.add_run(text)
            return new_para

        cursor = doc.paragraphs[pres_idx]
        if socios:
            for idx, s in enumerate(socios, start=1):
                nome = (s.get("nome") or "").strip() or "________________________"
                cursor = insert_paragraph_after(cursor, f"{idx}. {nome}")
        else:
            for idx in range(1,5):
                cursor = insert_paragraph_after(cursor, f"{idx}. ________________________")

    # 3) Rebuild signatures: delete everything after "ASSINATURAS DOS SÓCIOS:"
    sig_idx = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith("ASSINATURAS DOS SÓCIOS"):
            sig_idx = i
            break
    if sig_idx is not None:
        # remove all following paragraphs
        for p in list(doc.paragraphs)[sig_idx+1:]:
            _remove_paragraph(p)

        # add signatures
        last = doc.paragraphs[sig_idx]
        from docx.oxml import OxmlElement
        from docx.text.paragraph import Paragraph
        def insert_after(paragraph, text):
            new_p = OxmlElement("w:p")
            paragraph._p.addnext(new_p)
            new_para = Paragraph(new_p, paragraph._parent)
            if text:
                new_para.add_run(text)
            return new_para

        for s in socios:
            last = insert_after(last, "")
            last = insert_after(last, "_______________________________________")
            nome = (s.get("nome") or "").strip() or "________________________"
            cpf = (s.get("cpf") or "").strip()
            last = insert_after(last, nome)
            if cpf:
                last = insert_after(last, f"CPF nº {cpf}")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

def docx_to_pdf_via_libreoffice(docx_bytes: bytes) -> bytes | None:
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

def make_report_xlsx(rows: list[dict]) -> bytes:
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
            r.get("arquivo_origem",""),
            r.get("status",""),
            r.get("mensagem",""),
            "sim" if r.get("ocr_usado") else "não",
            r.get("metodo",""),
            r.get("razao_social",""),
            r.get("cnpj",""),
            r.get("nire",""),
            r.get("cidade_uf",""),
            r.get("qtd_socios",0),
            "sim" if r.get("docx_bytes") else "não",
            "sim" if r.get("pdf_bytes") else "não",
            r.get("tempo_ms",""),
        ])
    for col in range(1, len(headers)+1):
        ws.column_dimensions[get_column_letter(col)].width = 22
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()

def load_batch_files(uploaded_files, uploaded_zip):
    files = []
    if uploaded_files:
        for f in uploaded_files:
            files.append((f.name, f.read()))
    if uploaded_zip:
        zbytes = uploaded_zip.read()
        with zipfile.ZipFile(io.BytesIO(zbytes)) as z:
            for name in z.namelist():
                if name.lower().endswith(".pdf") and not name.endswith("/"):
                    files.append((Path(name).name, z.read(name)))
                if name.lower().endswith(".docx") and not name.endswith("/"):
                    files.append((Path(name).name, z.read(name)))
    return files

def process_one(name: str, bts: bytes, template_path: str, try_pdf: bool, use_gpt: bool) -> dict:
    t0 = time.time()
    res = {
        "arquivo_origem": name,
        "status": "",
        "mensagem": "",
        "ocr_usado": False,
        "metodo": "regex",
        "razao_social": "",
        "cnpj": "",
        "nire": "",
        "cidade_uf": "",
        "qtd_socios": 0,
        "tempo_ms": "",
        "docx_bytes": None,
        "pdf_bytes": None,
    }
    try:
        text = ""
        if name.lower().endswith(".pdf"):
            text, ocr_used = extract_text_from_pdf(bts)
            res["ocr_usado"] = ocr_used
        elif name.lower().endswith(".docx"):
            text = extract_text_from_docx(bts)
        else:
            res["status"] = "ERRO"
            res["mensagem"] = "Formato não suportado (use PDF/DOCX)."
            return res

        if not text.strip():
            res["status"] = "ERRO"
            res["mensagem"] = "Texto vazio (PDF pode ser imagem e OCR falhou/indisponível)."
            return res

        data = {}
        raw = ""
        if use_gpt:
            data, raw = gpt_extract(text)
            if data:
                res["metodo"] = "gpt"
            else:
                data = regex_extract(text)
                res["metodo"] = "regex (fallback)"
        else:
            data = regex_extract(text)

        # build ctx for template fill
        dia_str = str(ata_date.day)
        mes_str = PT_MONTHS[ata_date.month]
        ano_str = str(ata_date.year)
        dia_str = str(ata_date.day)
        mes_str = PT_MONTHS[ata_date.month]
        ano_str = str(ata_date.year)

        ctx = {
            **data,
            "presidente": DEFAULT_PRESIDENTE,
            "secretario": DEFAULT_ADVOGADO,
            "dia": dia_str,
            "mes": mes_str,
            "ano": ano_str,
        }

        [:-1] + f\'            "dia": dia_str,
            "mes": mes_str,
            "ano": ano_str,
        }\'
docx_bytes = fill_template_docx(template_path, ctx)
        res["docx_bytes"] = docx_bytes
        if try_pdf:
            res["pdf_bytes"] = docx_to_pdf_via_libreoffice(docx_bytes)

        res["razao_social"] = data.get("razao_social","")
        res["cnpj"] = data.get("cnpj","")
        res["nire"] = data.get("nire","")
        res["cidade_uf"] = data.get("cidade_uf","")
        res["qtd_socios"] = len(data.get("socios") or [])

        # basic quality gates
        issues = []
        if not res["razao_social"]:
            issues.append("Razão social ausente")
        if not res["cnpj"]:
            issues.append("CNPJ ausente")
        if res["qtd_socios"] == 0:
            issues.append("Sócios não encontrados")
        if issues:
            res["status"] = "PENDENTE"
            res["mensagem"] = "; ".join(issues)
        else:
            res["status"] = "OK"
            res["mensagem"] = "Gerado"

        return res
    except Exception as e:
        res["status"] = "ERRO"
        res["mensagem"] = f"{type(e).__name__}: {e}"
        return res
    finally:
        res["tempo_ms"] = int((time.time() - t0)*1000)

# ================= UI =================
st.set_page_config(page_title="Gerador de ATA (Template correto + GPT opcional)", layout="wide")
st.title("Gerador de ATA — Template correto + extração (Regex/GPT)")

with st.sidebar:
    st.subheader("Template")
    template_path = st.text_input("Caminho do template", value="templates/TEMPLATE_ATA.docx")
    try_pdf = st.checkbox("Gerar PDF (requer LibreOffice)", value=False)

    st.subheader("Padrões")
    st.text_input("Presidente (fixo)", value=DEFAULT_PRESIDENTE, disabled=True)
    st.text_input("Advogado (fixo)", value=DEFAULT_ADVOGADO, disabled=True)

    st.subheader("Data da ATA (padrão = hoje)")
    ata_date = st.date_input("Data", value=NOW.date())

    st.subheader("Extração")
    use_gpt = st.checkbox("Usar GPT para melhorar extração (recomendado)", value=False)
    st.caption("Se marcar, configure OPENAI_API_KEY em .streamlit/secrets.toml no Streamlit Cloud.")

tab1, tab2 = st.tabs(["Individual", "Lote"])

with tab1:
    st.subheader("Individual")
    up = st.file_uploader("Suba 1 contrato (PDF ou DOCX)", type=["pdf","docx"])
    if up:
        bts = up.read()
        res = process_one(up.name, bts, template_path, try_pdf, use_gpt)

        st.write(f"**Status:** {res['status']}  |  **Método:** {res['metodo']}")
        st.write(f"**Mensagem:** {res['mensagem']}")
        st.write(f"**Sócios encontrados:** {res.get('qtd_socios',0)}")

        base = Path(up.name).stem
        if res.get("docx_bytes"):
            st.download_button("Baixar ATA (DOCX)", data=res["docx_bytes"], file_name=f"ATA_{base}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        if res.get("pdf_bytes"):
            st.download_button("Baixar ATA (PDF)", data=res["pdf_bytes"], file_name=f"ATA_{base}.pdf", mime="application/pdf")

with tab2:
    st.subheader("Lote")
    uploaded_files = st.file_uploader("Vários arquivos (PDF/DOCX)", type=["pdf","docx"], accept_multiple_files=True)
    uploaded_zip = st.file_uploader("Ou ZIP com PDFs/DOCXs", type=["zip"])
    batch = load_batch_files(uploaded_files, uploaded_zip)
    st.write(f"Arquivos carregados: **{len(batch)}**")

    continue_on_error = st.checkbox("Continuar mesmo se houver erro", value=True)
    include_excel_inside_zip = st.checkbox("Incluir Excel dentro do ZIP", value=True)

    if st.button("Processar lote", disabled=(len(batch)==0)):
        rows = []
        prog = st.progress(0.0)
        box = st.empty()

        for i, (name, bts) in enumerate(batch, start=1):
            box.write(f"Processando {i}/{len(batch)}: **{name}**")
            r = process_one(name, bts, template_path, try_pdf, use_gpt)
            rows.append(r)
            prog.progress(i/len(batch))
            if (not continue_on_error) and r["status"] == "ERRO":
                st.error(f"Parado no erro: {name} — {r['mensagem']}")
                break

        excel_bytes = make_report_xlsx(rows)

        zip_out = io.BytesIO()
        with zipfile.ZipFile(zip_out, "w", zipfile.ZIP_DEFLATED) as z:
            for r in rows:
                base = Path(r["arquivo_origem"]).stem
                if r.get("docx_bytes"):
                    z.writestr(f"ATAs/{base}.docx", r["docx_bytes"])
                if r.get("pdf_bytes"):
                    z.writestr(f"ATAs/{base}.pdf", r["pdf_bytes"])
            if include_excel_inside_zip:
                z.writestr("relatorio_processamento.xlsx", excel_bytes)

        zip_out.seek(0)

        st.download_button("Baixar Relatório (Excel)", data=excel_bytes, file_name="relatorio_processamento.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button("Baixar ZIP (ATAs + Excel opcional)", data=zip_out.getvalue(), file_name="atas_geradas.zip", mime="application/zip")

        st.dataframe([{
            "arquivo": r["arquivo_origem"],
            "status": r["status"],
            "metodo": r["metodo"],
            "ocr": "sim" if r.get("ocr_usado") else "não",
            "socios": r.get("qtd_socios",0),
            "mensagem": r.get("mensagem","")
        } for r in rows])
