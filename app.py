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
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# OCR fallback
try:
    import pytesseract
    from pdf2image import convert_from_bytes
    OCR_OK = True
except Exception:
    OCR_OK = False

NOW = datetime.now()

CNPJ_RE = re.compile(r"\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b")
CPF_RE  = re.compile(r"\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b")
CEP_RE  = re.compile(r"\b\d{2}\.?\d{3}-?\d{3}\b")
NIRE_RE = re.compile(r"\bNIRE\s*(?:n[º°o]?\.?)?\s*[:\-]?\s*([0-9.\-]{5,})", re.IGNORECASE)

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

def find_first(regex: re.Pattern, text: str) -> str:
    m = regex.search(text or "")
    return m.group(0) if m else ""

def find_cnpj(text: str) -> str:
    return find_first(CNPJ_RE, text)

def find_nire(text: str) -> str:
    m = NIRE_RE.search(text or "")
    return m.group(1) if m else ""

def guess_company_name(text: str) -> str:
    if not text:
        return ""
    # patterns
    m = re.search(r"(?:(?:SOCIEDADE|EMPRESA)\s+)?([A-ZÁÂÃÉÊÍÓÔÕÚÇ0-9][A-ZÁÂÃÉÊÍÓÔÕÚÇ0-9\s\.\-&]+(?:LTDA|S\/A|SA|EIRELI|ME|EPP))", text)
    if m:
        return normalize_spaces(m.group(1))[:160]
    # clause: "gira sob o nome empresarial ..."
    m = re.search(r"gira\s+sob\s+o\s+nome\s+empresarial\s+(.+)", text, flags=re.IGNORECASE)
    if m:
        return normalize_spaces(m.group(1).split("\n")[0].strip(" .;:-"))[:160]
    return ""

def extract_address(text: str) -> tuple[str,str]:
    """
    Returns (endereco_completo, cidade_uf).
    Looks for "localizada no endereço ..." or "tem sede: ..."
    """
    if not text:
        return "", ""
    # Prefer explicit "localizada no endereço"
    m = re.search(r"localizad[ao]\s+no\s+endere[cç]o\s+(.+?)(?:\.\s|\n)", text, flags=re.IGNORECASE|re.DOTALL)
    if not m:
        m = re.search(r"tem\s+sede\s*:\s*(.+?)(?:\.\s|\n)", text, flags=re.IGNORECASE|re.DOTALL)
    addr = normalize_spaces(m.group(1)) if m else ""
    # City/UF conservative
    cityuf = ""
    m2 = re.search(r"\b([A-ZÁÂÃÉÊÍÓÔÕÚÇ][A-Za-zÁÂÃÉÊÍÓÔÕÚÇ\s]{2,})\s*/\s*([A-Z]{2})\b", addr)
    if m2:
        cityuf = f"{normalize_spaces(m2.group(1))}/{m2.group(2)}"
    else:
        m3 = re.search(r"\b([A-ZÁÂÃÉÊÍÓÔÕÚÇ][A-Za-zÁÂÃÉÊÍÓÔÕÚÇ\s]{2,})\s*[-]\s*([A-Z]{2})\b", addr)
        if m3:
            cityuf = f"{normalize_spaces(m3.group(1))}/{m3.group(2)}"
    return addr, cityuf

def extract_partners(text: str) -> list[dict]:
    """
    Robust extraction:
    1) Signature pairs: Name line then CPF line (common in atas)
    2) Inline patterns: "NOME, ..., CPF nº xxx"
    3) Presence list (numbered) if found
    Deduplicate by CPF.
    """
    if not text:
        return []

    lines = [normalize_spaces(l) for l in text.splitlines()]
    partners = []
    seen_cpfs = set()

    # 1) signature pairs
    for i in range(1, len(lines)):
        line = lines[i]
        if re.search(r"\bCPF\b", line, flags=re.IGNORECASE):
            mcpf = CPF_RE.search(line)
            if not mcpf:
                continue
            cpf = mcpf.group(0)
            name = lines[i-1]
            if name and not re.search(r"\d", name) and len(name) >= 5:
                if cpf not in seen_cpfs:
                    partners.append({"nome": name, "cpf": cpf})
                    seen_cpfs.add(cpf)

    # 2) inline patterns
    inline = re.finditer(r"([A-ZÁÂÃÉÊÍÓÔÕÚÇ][A-ZÁÂÃÉÊÍÓÔÕÚÇ\s]{4,120})\s*,[^.\n]{0,120}?\bCPF\s*(?:n[º°o]?\.?)?\s*[:\-]?\s*("+CPF_RE.pattern+r")", text)
    for m in inline:
        name = normalize_spaces(m.group(1))
        cpf = m.group(2)
        if cpf not in seen_cpfs:
            partners.append({"nome": name, "cpf": cpf})
            seen_cpfs.add(cpf)

    # 3) presence numbered list block
    if not partners:
        # try to capture after "DA PRESENÇA" until "Os quais" or "DA COMPOSIÇÃO"
        mblock = re.search(r"DA\s+PRESEN[ÇC]A(.+?)(?:DA\s+COMPOSI[ÇC][AÃ]O|Os\s+quais|Portanto)", text, flags=re.IGNORECASE|re.DOTALL)
        block = mblock.group(1) if mblock else ""
        for m in re.finditer(r"^\s*\d+\.\s*([A-ZÁÂÃÉÊÍÓÔÕÚÇ][A-ZÁÂÃÉÊÍÓÔÕÚÇ\s]{4,120})\s*,?\s*$", block, flags=re.MULTILINE):
            partners.append({"nome": normalize_spaces(m.group(1)), "cpf": ""})

    # fallback: at least collect CPFs
    if not partners:
        cpfs = []
        for m in CPF_RE.finditer(text):
            c = m.group(0)
            if c not in cpfs:
                cpfs.append(c)
        for c in cpfs:
            partners.append({"nome": "", "cpf": c})

    return partners

def validate(ctx: dict) -> list[str]:
    pend = []
    if not ctx.get("RAZAO_SOCIAL"):
        pend.append("Razão social ausente")
    if not ctx.get("CNPJ"):
        pend.append("CNPJ ausente")
    if not ctx.get("ENDERECO"):
        pend.append("Endereço/sede ausente")
    if not ctx.get("CIDADE_UF"):
        pend.append("Cidade/UF ausente")
    socios = ctx.get("SOCIOS") or []
    if not socios:
        pend.append("Sócios ausentes")
    return pend

def make_ata_docx(ctx: dict) -> bytes:
    """
    Generates a DOCX modeled after the user's example PDF structure.
    """
    doc = Document()

    # Base font size
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    # Title
    p = doc.add_paragraph("ATA DE SÓCIOS")
    p.runs[0].bold = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Company header paragraph
    header = (
        f"Sociedade empresária limitada – {ctx.get('RAZAO_SOCIAL','').strip() or '________________________'}"
        f", inscrita no CNPJ/MF nº. {ctx.get('CNPJ','').strip() or '____________________'}"
    )
    if ctx.get("NIRE"):
        header += f" – NIRE: {ctx.get('NIRE')}"
    header += f" localizada no endereço {ctx.get('ENDERECO','').strip() or '__________________________________________'}."
    doc.add_paragraph(header)

    # Legal paragraph
    doc.add_paragraph(
        "A presente REUNIÃO DE SÓCIOS foi realizada em obediência à legislação civil, "
        "notadamente, mas não se limitando a estes, aos artigos 1.010, 1.071, 1.072, 1.073, 1.074, "
        "1.075, 1.076, 1.078 e 1.079, todos do Código Civil Brasileiro e em observância a Lei de "
        "Sociedades Anônimas, Lei 6.404/76 no que for pertinente, bem como as disposições societárias "
        "do Contrato Social da Sociedade ou eventual Acordo de Sócios existente."
    )

    # Meeting opening paragraph
    doc.add_paragraph(
        f"Aos dias {ctx.get('DIA_EXT','___')} do mês de {ctx.get('MES_EXT','__________')} do ano de {ctx.get('ANO','____')}, "
        f"às {ctx.get('HORA','__:__')} horas na sede da {ctx.get('RAZAO_SOCIAL','________________________')}, "
        f"localizada no endereço {ctx.get('ENDERECO','__________________________________________')}."
    )

    # Presence
    doc.add_paragraph("DA PRESENÇA - Foi realizada REUNIÃO DE SÓCIOS, na qual compareceram, em primeira convocação, os seguintes sócios:")
    for i, s in enumerate(ctx.get("SOCIOS", []), start=1):
        nome = (s.get("nome") or "").strip() or "________________________"
        doc.add_paragraph(f"{i}. {nome}", style=None)

    doc.add_paragraph("Os quais integralizam conjuntamente capital social de 100% da sociedade.")
    doc.add_paragraph("Portanto, foi alcançado quórum para se efetivar esta REUNIÃO DE SÓCIOS.")

    # Mesa
    presidente = ctx.get("PRESIDENTE") or "________________________"
    secretario = ctx.get("SECRETARIO") or "________________________"
    doc.add_paragraph(
        f"DA COMPOSIÇÃO DA MESA – Em cumprimento à legislação, a Assembleia foi presidida por {presidente} "
        f"e pelo SECRETÁRIO {secretario}, escolhidos entre os presentes."
    )

    # Publicações
    doc.add_paragraph("DAS PUBLICAÇÕES – Convocação realizada em formato DIGITAL.")

    # Ordem do dia (fixa, alinhada ao seu exemplo)
    doc.add_paragraph("DA ORDEM DO DIA – Esta REUNIÃO OU ASSEMBLEIA DE SÓCIOS teve como ordem do dia:")
    itens = [
        "Aprovação da distribuição dos lucros e dividendos acumulados até o exercício de 2025;",
        "Definição do cronograma de pagamentos no período de 2026 a 2028;",
        "Especificações de eventualidades e ajustes relacionados às condições financeiras da empresa e mudanças tributárias."
    ]
    for i, it in enumerate(itens, start=1):
        doc.add_paragraph(f"{i}. {it}")

    # Deliberações (texto base)
    doc.add_paragraph("DAS DELIBERAÇÕES – Iniciada a Assembleia, procedeu-se a leitura da Ata passada e verificou-se que inexiste pendências a serem incluídas na Ordem do dia, motivo pelo qual foi dado prosseguimento às deliberações pautadas.")

    doc.add_paragraph("I. APROVAÇÃO DA DISTRIBUIÇÃO DOS LUCROS E DIVIDENDOS")
    doc.add_paragraph(
        "Os sócios, por unanimidade, aprovaram a distribuição dos lucros e dividendos acumulados até o exercício de 2025, "
        "conforme apurado nos demonstrativos contábeis aprovados em Assembleia Geral e com base nos balanços patrimoniais da empresa."
    )
    doc.add_paragraph(
        "A partir da data de assinatura desta ata e até 31 de dezembro de 2025, todo o lucro acumulado e que for aferido por meio das atividades "
        "da empresa será objeto de distribuição intermediária entre os sócios, que será pago proporcionalmente à participação societária de cada sócio "
        "no capital da empresa. Este montante inclui os lucros acumulados até o exercício fiscal de 2025 e será pago em consonância com o PL 1087/2025 "
        "ou nos termos da Lei que restar definitivamente aprovada e vigente, que autoriza o pagamento até 31 de dezembro de 2028."
    )

    doc.add_paragraph("II. DEFINIÇÃO DO CRONOGRAMA DE PAGAMENTO")
    doc.add_paragraph(
        "Fica definido, por unanimidade entre as partes, que o pagamento dos lucros e dividendos acumulados referentes ao período anterior será realizado "
        "no intervalo compreendido entre os anos de 2026 a 2028. O referido pagamento será efetuado de forma intermediária e mensal."
    )

    doc.add_paragraph("III. EVENTUALIDADES, AJUSTES E CONDICIONANTES")
    doc.add_paragraph(
        "Os sócios deliberaram que:\n"
        "a) Ajustes no cronograma de pagamento: O cronograma de pagamento poderá ser livremente ajustado pela administração da empresa conforme disponibilidade de caixa, "
        "condições financeiras da empresa e eventuais alterações tributárias ou de mercado que impactem a saúde financeira do negócio.\n"
        "b) Eventos de força maior ou caso fortuito: Caso, por qualquer motivo justificado, como falta de caixa, mudanças extraordinárias de impostos, força maior ou caso fortuito, "
        "os valores deliberados não possam ser efetivamente distribuídos até a data limite de 31 de dezembro de 2028, qualquer saldo residual será objeto de nova deliberação pelos sócios. "
        "Essa nova deliberação deverá priorizar a análise das condições financeiras e tributárias da empresa, buscando soluções para o pagamento ou replanejamento de forma a minimizar impactos financeiros e fiscais.\n"
        "c) Impactos tributários: A distribuição do saldo, quando aplicável, será condicionada à avaliação dos impactos tributários das normas vigentes e à preservação da saúde financeira e operacional da empresa."
    )

    doc.add_paragraph(
        "DO ENCERRAMENTO E APROVAÇÃO DA ATA - Por fim, a palavra foi concedida à aquele que dela quisesse fazer uso para discorrer sobre os assuntos de interesse social. "
        "Não existindo manifestações, o PRESIDENTE encerrou a REUNIÃO. O SECRETÁRIO lavrou a presente ata, aprovada, executou a sua leitura, e ela foi assinada pelos sócios presentes, "
        "pelo SECRETÁRIO e pelo PRESIDENTE. Sem mais para o presente."
    )

    # Date line
    doc.add_paragraph(f"{(ctx.get('CIDADE_UF') or '________________')}, {ctx.get('DIA_EXT_MAIUS','___')} de {ctx.get('MES_EXT_MAIUS','__________')} de {ctx.get('ANO','____')}.")

    # Signatures
    doc.add_paragraph("ASSINATURAS DOS SÓCIOS:")
    for s in ctx.get("SOCIOS", []):
        doc.add_paragraph("_______________________________________")
        nome = (s.get("nome") or "").strip() or "________________________"
        cpf = (s.get("cpf") or "").strip()
        doc.add_paragraph(nome)
        if cpf:
            doc.add_paragraph(f"CPF nº {cpf}")

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
        "arquivo_origem","status","mensagem","ocr_usado",
        "razao_social","cnpj","nire","cidade_uf","qtd_socios",
        "pendencias","docx_gerado","pdf_gerado","tempo_ms"
    ]
    ws.append(headers)
    for r in rows:
        ws.append([
            r.get("arquivo_origem",""),
            r.get("status",""),
            r.get("mensagem",""),
            "sim" if r.get("ocr_usado") else "não",
            r.get("razao_social",""),
            r.get("cnpj",""),
            r.get("nire",""),
            r.get("cidade_uf",""),
            r.get("qtd_socios",0),
            r.get("pendencias",""),
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
    return files

def process_pdf(file_name: str, file_bytes: bytes, try_pdf: bool, defaults: dict) -> dict:
    t0 = time.time()
    res = {
        "arquivo_origem": file_name,
        "status": "",
        "mensagem": "",
        "ocr_usado": False,
        "razao_social": "",
        "cnpj": "",
        "nire": "",
        "cidade_uf": "",
        "qtd_socios": 0,
        "pendencias": "",
        "tempo_ms": "",
        "docx_bytes": None,
        "pdf_bytes": None,
    }
    try:
        text, ocr_used = extract_text_from_pdf(file_bytes)
        res["ocr_usado"] = ocr_used
        if not text.strip():
            res["status"] = "ERRO"
            res["mensagem"] = "Não foi possível extrair texto (PDF escaneado e OCR indisponível/fracassou)."
            return res

        razao = guess_company_name(text)
        cnpj = find_cnpj(text)
        nire = find_nire(text)
        endereco, cityuf = extract_address(text)
        socios = extract_partners(text)

        ctx = {
            "RAZAO_SOCIAL": razao,
            "CNPJ": cnpj,
            "NIRE": nire,
            "ENDERECO": endereco,
            "CIDADE_UF": cityuf or defaults.get("cidade_uf",""),
            "SOCIOS": socios,
            "PRESIDENTE": defaults.get("presidente",""),
            "SECRETARIO": defaults.get("secretario",""),
            "DIA_EXT": defaults.get("dia_ext","___"),
            "MES_EXT": defaults.get("mes_ext","__________"),
            "ANO": str(defaults.get("ano", NOW.year)),
            "HORA": defaults.get("hora","__:__"),
            "DIA_EXT_MAIUS": defaults.get("dia_ext_maius","___"),
            "MES_EXT_MAIUS": defaults.get("mes_ext_maius","__________"),
        }

        pend = validate(ctx)
        docx_bytes = make_ata_docx(ctx)
        res["docx_bytes"] = docx_bytes
        if try_pdf:
            res["pdf_bytes"] = docx_to_pdf_via_libreoffice(docx_bytes)

        res["razao_social"] = razao
        res["cnpj"] = cnpj
        res["nire"] = nire
        res["cidade_uf"] = ctx["CIDADE_UF"]
        res["qtd_socios"] = len(socios)
        res["pendencias"] = "; ".join(pend)

        if pend:
            res["status"] = "PENDENTE"
            res["mensagem"] = "Gerado com pendências (ver relatório)."
        else:
            res["status"] = "OK"
            res["mensagem"] = "Gerado com sucesso."
        if res["ocr_usado"]:
            res["mensagem"] += " (OCR aplicado)"
        return res
    except Exception as e:
        res["status"] = "ERRO"
        res["mensagem"] = f"Exceção: {type(e).__name__}: {e}"
        return res
    finally:
        res["tempo_ms"] = int((time.time() - t0) * 1000)

# ================= UI =================
st.set_page_config(page_title="Gerador de ATA (PL1087/2025)", layout="wide")
st.title("Gerador de ATA — PL 1087/2025 (Individual e Lote)")
st.caption("Gera ATA no formato do seu exemplo (ATA DE SÓCIOS). Extrai dados do PDF (com OCR automático se precisar), gera DOCX e relatório Excel no lote.")

with st.sidebar:
    st.subheader("Padrões da ATA (ajustáveis)")
    presidente = st.text_input("Presidente (ex.: Contador ...)", value="")
    secretario = st.text_input("Secretário", value="")
    hora = st.text_input("Hora (ex.: 08:05)", value="__:__")
    dia_ext = st.text_input("Dia por extenso (ex.: dois)", value="___")
    mes_ext = st.text_input("Mês por extenso (ex.: dezembro)", value="__________")
    ano = st.number_input("Ano", min_value=2000, max_value=2100, value=NOW.year, step=1)
    st.markdown("---")
    cidade_fallback = st.text_input("Cidade/UF fallback (se não achar no PDF)", value="")

    try_pdf = st.checkbox("Tentar gerar PDF (LibreOffice no servidor)", value=False)
    st.caption("No Streamlit Cloud, mantenha `packages.txt` com `libreoffice` para o PDF.")

defaults = {
    "presidente": presidente,
    "secretario": secretario,
    "hora": hora,
    "dia_ext": dia_ext,
    "mes_ext": mes_ext,
    "ano": int(ano),
    "dia_ext_maius": dia_ext.upper() if dia_ext else "___",
    "mes_ext_maius": mes_ext.upper() if mes_ext else "__________",
    "cidade_uf": cidade_fallback,
}

tab1, tab2 = st.tabs(["Individual", "Lote"])

with tab1:
    st.subheader("Individual")
    up = st.file_uploader("Suba 1 PDF do contrato/alteração", type=["pdf"], key="onepdf")
    if up:
        b = up.read()
        res = process_pdf(up.name, b, try_pdf=try_pdf, defaults=defaults)

        st.write(f"**Status:** {res['status']}")
        st.write(f"**Mensagem:** {res['mensagem']}")
        if res.get("pendencias"):
            st.warning(f"Pendências: {res['pendencias']}")
        st.write(f"**Sócios encontrados:** {res.get('qtd_socios',0)}")

        base = Path(up.name).stem
        if res.get("docx_bytes"):
            st.download_button("Baixar ATA (DOCX)", data=res["docx_bytes"], file_name=f"ATA_{base}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        if res.get("pdf_bytes"):
            st.download_button("Baixar ATA (PDF)", data=res["pdf_bytes"], file_name=f"ATA_{base}.pdf", mime="application/pdf")

with tab2:
    st.subheader("Lote")
    st.write("Envie vários PDFs ou um ZIP com PDFs. Saída: ZIP com DOCX/PDF + relatório Excel.")

    uploaded_files = st.file_uploader("Vários PDFs", type=["pdf"], accept_multiple_files=True, key="many")
    uploaded_zip = st.file_uploader("ZIP com PDFs", type=["zip"], key="zip")

    batch = load_batch_files(uploaded_files, uploaded_zip)
    st.write(f"Arquivos carregados: **{len(batch)}**")

    continue_on_error = st.checkbox("Continuar mesmo se houver erro", value=True)
    include_excel_inside_zip = st.checkbox("Incluir Excel dentro do ZIP", value=True)

    if st.button("Processar lote", disabled=(len(batch) == 0)):
        rows = []
        prog = st.progress(0.0)
        box = st.empty()

        for i, (name, bts) in enumerate(batch, start=1):
            box.write(f"Processando {i}/{len(batch)}: **{name}**")
            r = process_pdf(name, bts, try_pdf=try_pdf, defaults=defaults)
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

        ok = sum(1 for r in rows if r["status"] == "OK")
        pend = sum(1 for r in rows if r["status"] == "PENDENTE")
        err = sum(1 for r in rows if r["status"] == "ERRO")
        st.success(f"Concluído. OK: {ok} | PENDENTE: {pend} | ERRO: {err}")

        st.download_button("Baixar Relatório (Excel)", data=excel_bytes, file_name="relatorio_processamento.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button("Baixar ZIP (ATAs + Excel opcional)", data=zip_out.getvalue(), file_name="atas_geradas.zip", mime="application/zip")

        st.dataframe([{
            "arquivo": r["arquivo_origem"],
            "status": r["status"],
            "ocr": "sim" if r.get("ocr_usado") else "não",
            "socios": r.get("qtd_socios",0),
            "pendencias": r.get("pendencias",""),
        } for r in rows])
