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
from docx import Document

# Optional: better template rendering for {{VAR}} placeholders
try:
    from docxtpl import DocxTemplate
    DOCXTPL_OK = True
except Exception:
    DOCXTPL_OK = False

# PDF native text extraction
try:
    import pdfplumber
    PDFPLUMBER_OK = True
except Exception:
    PDFPLUMBER_OK = False

# OCR fallback (scanned PDFs)
try:
    import pytesseract
    from pdf2image import convert_from_bytes
    OCR_OK = True
except Exception:
    OCR_OK = False

# Excel report
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


# =========================
# Config
# =========================
TEMPLATE_PATH_DEFAULT = "templates/MODELO_ATA.docx"
DATE_NOW = datetime.now()

CNPJ_RE = re.compile(r"\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b")
CPF_RE  = re.compile(r"\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b")


def normalize_spaces(s: str) -> str:
    s = (s or "").replace("\u00A0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()


def extract_text_from_pdf(file_bytes: bytes) -> tuple[str, bool]:
    """
    Returns: (text, ocr_used)
    Strategy:
      1) Try native extraction (pdfplumber)
      2) If too short, fallback to OCR (pdf2image + tesseract)
    """
    text_native = ""
    if PDFPLUMBER_OK:
        try:
            parts = []
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                for page in pdf.pages:
                    t = (page.extract_text() or "").strip()
                    if t:
                        parts.append(t)
            text_native = "\n".join(parts).strip()
        except Exception:
            text_native = ""

    # If native extraction is good enough, return it
    if len(text_native) >= 200:
        return text_native, False

    # OCR fallback
    if not OCR_OK:
        return text_native, False

    try:
        ocr_parts = []
        images = convert_from_bytes(file_bytes, dpi=250)  # 200–300 is a good range
        for img in images:
            ocr_text = pytesseract.image_to_string(
                img,
                lang="por",
                config="--oem 3 --psm 6"
            )
            ocr_text = (ocr_text or "").strip()
            if ocr_text:
                ocr_parts.append(ocr_text)
        text_ocr = "\n".join(ocr_parts).strip()

        # Return whichever is longer (usually OCR)
        if len(text_ocr) > len(text_native):
            return text_ocr, True
        return text_native, False
    except Exception:
        return text_native, False


def extract_text_from_docx(file_bytes: bytes) -> str:
    doc = Document(io.BytesIO(file_bytes))
    parts = []
    for p in doc.paragraphs:
        if p.text and p.text.strip():
            parts.append(p.text)
    return "\n".join(parts)


def find_cnpj(text: str) -> str:
    m = CNPJ_RE.search(text or "")
    return m.group(0) if m else ""


def find_cpfs(text: str) -> list[str]:
    if not text:
        return []
    cpfs = [m.group(0) for m in CPF_RE.finditer(text)]
    seen = set()
    out = []
    for c in cpfs:
        if c not in seen:
            seen.add(c)
            out.append(c)
    return out


def guess_company_name(text: str) -> str:
    if not text:
        return ""
    patterns = [
        r"(?:DENOMINAÇÃO|RAZÃO SOCIAL|NOME EMPRESARIAL)\s*[:\-]\s*(.+)",
        r"(?:EMPRESA|SOCIEDADE)\s*[:\-]\s*(.+)",
    ]
    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            line = m.group(1).split("\n")[0].strip(" .;:-")
            return normalize_spaces(line)[:140]

    # Fallback: first strong uppercase line
    candidates = []
    for line in text.splitlines():
        l = normalize_spaces(line)
        if len(l) >= 18:
            upper_ratio = sum(ch.isupper() for ch in l) / max(len(l), 1)
            if upper_ratio > 0.7 and not re.search(r"\b(CLÁUSULA|ALTERA(Ç|C)ÃO|CONTRATO|CNPJ|CPF)\b", l, re.IGNORECASE):
                candidates.append(l)
    return candidates[0][:140] if candidates else ""


def guess_partners(text: str) -> list[dict]:
    partners = []
    if not text:
        return partners

    lines = [normalize_spaces(l) for l in text.splitlines() if normalize_spaces(l)]
    for i, line in enumerate(lines):
        cpf_m = CPF_RE.search(line)
        if not cpf_m:
            continue
        cpf = cpf_m.group(0)
        before = line[:cpf_m.start()].strip(" ,;-:")
        name = ""
        if len(before) >= 6:
            name = before
        elif i > 0 and len(lines[i - 1]) >= 6:
            name = lines[i - 1]

        name = re.sub(
            r"\b(CPF|RG|NACIONALIDADE|ESTADO CIVIL|PROFISS(Ã|A)O|RESIDENTE|DOMICILIAD[AO])\b.*",
            "",
            name,
            flags=re.IGNORECASE
        ).strip(" ,;-:")

        if any(p["cpf"] == cpf for p in partners):
            continue
        partners.append({"nome": name, "cpf": cpf})

    if not partners:
        for cpf in find_cpfs(text):
            partners.append({"nome": "", "cpf": cpf})

    return partners


def guess_city_uf(text: str) -> str:
    # Conservative: only return if explicit "Cidade - UF" or "Cidade/UF"
    if not text:
        return ""
    m = re.search(r"\b([A-ZÁÂÃÉÊÍÓÔÕÚÇ][A-Za-zÁÂÃÉÊÍÓÔÕÚÇ\s]{2,})\s*[-/]\s*([A-Z]{2})\b", text)
    if m:
        return f"{normalize_spaces(m.group(1))}/{m.group(2)}"
    return ""


def docx_contains_placeholders(docx_path: str) -> bool:
    try:
        doc = Document(docx_path)
        text = "\n".join(p.text for p in doc.paragraphs if p.text)
        return "{{" in text and "}}" in text
    except Exception:
        return False


def replace_in_doc(doc: Document, mapping: dict[str, str]) -> None:
    def replace_par(par):
        if not par.runs:
            return
        full = "".join(run.text for run in par.runs)
        new = full
        for k, v in mapping.items():
            if k in new:
                new = new.replace(k, v)
        if new != full:
            for run in par.runs:
                run.text = ""
            if par.runs:
                par.runs[0].text = new
            else:
                par.add_run(new)

    for p in doc.paragraphs:
        replace_par(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_par(p)


def build_signature_text(partners: list[dict]) -> str:
    chunks = []
    for p in partners:
        nome = (p.get("nome") or "").strip() or "Nome do Sócio"
        cpf = (p.get("cpf") or "").strip() or "___________________________"
        cpf_line = cpf if cpf.lower().startswith("cpf") else f"CPF: {cpf}"
        chunks.append(f"{nome}\n{cpf_line}\n")
    return "\n".join(chunks).strip()


def render_docx_from_template(template_path: str, context_ata: dict, legacy_mapping: dict) -> bytes:
    # Preferred: docxtpl + {{VAR}} placeholders
    if DOCXTPL_OK and docx_contains_placeholders(template_path):
        tpl = DocxTemplate(template_path)
        tpl.render(context_ata)
        out = io.BytesIO()
        tpl.save(out)
        return out.getvalue()

    # Legacy: replace XXXXX/___ patterns
    doc = Document(template_path)
    replace_in_doc(doc, legacy_mapping)

    # Append signatures at end (safer than trying to edit existing signature blocks)
    doc.add_paragraph("")
    doc.add_paragraph(build_signature_text(context_ata.get("SOCIOS", [])))

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


def validate_minimum(ctx: dict) -> list[str]:
    pend = []
    if not (ctx.get("SOCIEDADE") or "").strip():
        pend.append("Razão social ausente")
    if not (ctx.get("CNPJ") or "").strip():
        pend.append("CNPJ ausente")
    socios = ctx.get("SOCIOS") or []
    if not socios:
        pend.append("Lista de sócios ausente")
    else:
        if not any((s.get("cpf") or s.get("CPF") or "").strip() for s in socios):
            pend.append("Nenhum CPF de sócio identificado")
    if not (ctx.get("CIDADE_UF") or "").strip():
        pend.append("Cidade/UF ausente (preencher manualmente ou ajustar extração)")
    return pend


def make_report_xlsx(rows: list[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório"

    headers = [
        "arquivo_origem", "status", "mensagem",
        "ocr_usado",
        "razao_social", "cnpj", "qtd_socios",
        "pendencias",
        "docx_gerado", "pdf_gerado",
        "tempo_ms"
    ]
    ws.append(headers)

    for r in rows:
        ws.append([
            r.get("arquivo_origem", ""),
            r.get("status", ""),
            r.get("mensagem", ""),
            "sim" if r.get("ocr_usado") else "não",
            r.get("razao_social", ""),
            r.get("cnpj", ""),
            r.get("qtd_socios", 0),
            r.get("pendencias", ""),
            "sim" if r.get("docx_bytes") else "não",
            "sim" if r.get("pdf_bytes") else "não",
            r.get("tempo_ms", ""),
        ])

    for col in range(1, len(headers) + 1):
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
    return files


def process_contract_bytes(
    file_name: str,
    file_bytes: bytes,
    template_path: str,
    try_pdf: bool,
    default_city_uf: str,
    valores: dict
) -> dict:
    t0 = time.time()
    result = {
        "arquivo_origem": file_name,
        "status": "",
        "mensagem": "",
        "ocr_usado": False,
        "razao_social": "",
        "cnpj": "",
        "qtd_socios": 0,
        "pendencias": "",
        "tempo_ms": "",
        "docx_bytes": None,
        "pdf_bytes": None,
    }

    try:
        # 1) extract text
        if file_name.lower().endswith(".pdf"):
            text, ocr_used = extract_text_from_pdf(file_bytes)
            result["ocr_usado"] = bool(ocr_used)
        elif file_name.lower().endswith(".docx"):
            text = extract_text_from_docx(file_bytes)
        else:
            text = ""

        if not text.strip():
            result["status"] = "ERRO"
            result["mensagem"] = "Não foi possível extrair texto (PDF pode estar escaneado/imagem; OCR indisponível?)."
            return result

        text_norm = normalize_spaces(text)

        # 2) heuristics
        sociedade = guess_company_name(text)
        cnpj = find_cnpj(text_norm)
        socios = guess_partners(text)
        cidade_uf = guess_city_uf(text) or default_city_uf

        # 3) ATA context
        ctx = {
            "SOCIEDADE": sociedade,
            "CNPJ": cnpj,
            "DIA": str(DATE_NOW.day),
            "MES": "____________",
            "ANO": str(DATE_NOW.year),
            "HORA": "___",
            "CIDADE_UF": cidade_uf,
            "VALOR_LUCROS_ACUMULADOS": valores.get("lucros_acumulados", ""),
            "VALOR_LUCRO_EXERCICIO": valores.get("lucro_exercicio", ""),
            "SOCIOS": socios,
            "ASSINATURAS": build_signature_text(socios),
        }

        pend = validate_minimum(ctx)

        # 4) legacy mapping (your current template)
        legacy_mapping = {
            "XXXXXXXXXXXXX": ctx["SOCIEDADE"] or "______________________________",
            "____________________": ctx["CNPJ"] or "____________________",
            "Aos ___ dias": f"Aos {ctx['DIA']} dias",
            "mês de __________ de 2025": f"mês de {ctx['MES']} de {ctx['ANO']}",
            "às ___ horas": f"às {ctx['HORA']} horas",
            "Goiânia, dia ___ de _______ de 2025.": f"{ctx['CIDADE_UF']}, dia {ctx['DIA']} de {ctx['MES']} de {ctx['ANO']}.",
        }

        if ctx["VALOR_LUCROS_ACUMULADOS"]:
            legacy_mapping["R$ ____________"] = f"R$ {ctx['VALOR_LUCROS_ACUMULADOS']}"
        if ctx["VALOR_LUCRO_EXERCICIO"]:
            legacy_mapping["R$ ____________,"] = f"R$ {ctx['VALOR_LUCRO_EXERCICIO']},"

        # 5) DOCX
        docx_bytes = render_docx_from_template(template_path, ctx, legacy_mapping)
        result["docx_bytes"] = docx_bytes

        # 6) PDF optional
        if try_pdf:
            result["pdf_bytes"] = docx_to_pdf_via_libreoffice(docx_bytes)

        # 7) status + summary
        result["razao_social"] = sociedade
        result["cnpj"] = cnpj
        result["qtd_socios"] = len(socios)
        result["pendencias"] = "; ".join(pend)

        if pend:
            result["status"] = "PENDENTE"
            msg = "Gerado com campos faltantes (revisar pendências)."
        else:
            result["status"] = "OK"
            msg = "Gerado com sucesso."

        if result["ocr_usado"]:
            msg += " OCR aplicado."
        result["mensagem"] = msg

        return result

    except Exception as e:
        result["status"] = "ERRO"
        result["mensagem"] = f"Exceção: {type(e).__name__}: {e}"
        return result
    finally:
        result["tempo_ms"] = int((time.time() - t0) * 1000)


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Gerador de Atas (Individual e Lote)", layout="wide")
st.title("Gerador de ATA — Distribuição de Lucros / Dividendos")
st.caption("Gera DOCX (e opcional PDF) a partir do contrato social/última alteração. Inclui OCR para PDFs escaneados, modo em lote e relatório Excel.")

with st.sidebar:
    st.subheader("Configurações")
    template_path = st.text_input("Caminho do template DOCX", value=TEMPLATE_PATH_DEFAULT)
    try_pdf = st.checkbox("Tentar gerar PDF (LibreOffice/soffice no servidor)", value=False)
    default_city_uf = st.text_input("Cidade/UF padrão (quando não encontrado no contrato)", value="Goiânia/GO")

    st.markdown("**Valores (aplicados em todos no lote, se você quiser):**")
    lucros_acumulados = st.text_input("Lucros acumulados (ex.: 100.000,00)", value="")
    lucro_exercicio = st.text_input("Lucro do exercício (ex.: 50.000,00)", value="")

    st.markdown("---")
    st.caption("OCR ativo se: PDF não tiver texto suficiente e o servidor tiver Tesseract + Poppler instalados.")

tab1, tab2 = st.tabs(["Individual", "Lote"])

# Individual
with tab1:
    st.subheader("Geração Individual")
    contrato = st.file_uploader("Suba 1 contrato (PDF ou DOCX)", type=["pdf", "docx"], key="single")

    if contrato is not None:
        file_bytes = contrato.read()
        res = process_contract_bytes(
            file_name=contrato.name,
            file_bytes=file_bytes,
            template_path=template_path,
            try_pdf=try_pdf,
            default_city_uf=default_city_uf,
            valores={"lucros_acumulados": lucros_acumulados, "lucro_exercicio": lucro_exercicio},
        )

        st.write(f"**Status:** {res['status']}")
        st.write(f"**Mensagem:** {res['mensagem']}")
        if res.get("pendencias"):
            st.warning(f"Pendências: {res['pendencias']}")
        if contrato.name.lower().endswith(".pdf"):
            st.write(f"**OCR usado:** {'sim' if res.get('ocr_usado') else 'não'}")

        base = Path(contrato.name).stem
        if res.get("docx_bytes"):
            st.download_button(
                "Baixar ATA (DOCX)",
                data=res["docx_bytes"],
                file_name=f"ATA_{base}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        if res.get("pdf_bytes"):
            st.download_button(
                "Baixar ATA (PDF)",
                data=res["pdf_bytes"],
                file_name=f"ATA_{base}.pdf",
                mime="application/pdf"
            )

# Batch
with tab2:
    st.subheader("Geração em Lote")
    st.write("Envie **vários PDFs** ou **um ZIP** com PDFs. O sistema gera as atas e um relatório Excel.")

    uploaded_files = st.file_uploader(
        "Suba vários contratos (PDF)", type=["pdf"], accept_multiple_files=True, key="multi_pdfs"
    )
    uploaded_zip = st.file_uploader(
        "Ou envie um ZIP contendo PDFs", type=["zip"], key="zip_pdfs"
    )

    batch = load_batch_files(uploaded_files, uploaded_zip)
    st.write(f"Arquivos carregados: **{len(batch)}**")

    colA, colB = st.columns(2)
    with colA:
        continue_on_error = st.checkbox("Continuar mesmo se houver erro em algum arquivo", value=True)
    with colB:
        include_excel_inside_zip = st.checkbox("Incluir o Excel dentro do ZIP final", value=True)

    if st.button("Processar lote", disabled=(len(batch) == 0)):
        rows = []
        progress = st.progress(0)
        status_box = st.empty()

        for i, (name, bts) in enumerate(batch, start=1):
            status_box.write(f"Processando {i}/{len(batch)}: **{name}**")
            res = process_contract_bytes(
                file_name=name,
                file_bytes=bts,
                template_path=template_path,
                try_pdf=try_pdf,
                default_city_uf=default_city_uf,
                valores={"lucros_acumulados": lucros_acumulados, "lucro_exercicio": lucro_exercicio},
            )
            rows.append(res)
            progress.progress(i / len(batch))

            if (not continue_on_error) and res["status"] == "ERRO":
                st.error(f"Parado no erro: {name} — {res['mensagem']}")
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

        ok = sum(1 for r in rows if r["status"] == "OK")
        pend = sum(1 for r in rows if r["status"] == "PENDENTE")
        err = sum(1 for r in rows if r["status"] == "ERRO")
        st.success(f"Concluído. OK: {ok} | PENDENTE: {pend} | ERRO: {err}")

        st.download_button(
            "Baixar Relatório (Excel)",
            data=excel_bytes,
            file_name="relatorio_processamento.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.download_button(
            "Baixar ZIP (ATAs + opcional Excel)",
            data=zip_out.getvalue(),
            file_name="atas_geradas.zip",
            mime="application/zip",
        )

        st.dataframe([
            {
                "arquivo": r["arquivo_origem"],
                "status": r["status"],
                "ocr": "sim" if r.get("ocr_usado") else "não",
                "mensagem": r["mensagem"],
                "pendencias": r["pendencias"],
            } for r in rows
        ])
