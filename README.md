# Gerador de ATA (Streamlit)

App em Streamlit para gerar **Atas de Deliberação** a partir de contratos sociais/alterações (PDF/DOCX),
com suporte a:
- **Modo individual**
- **Geração em lote** (múltiplos PDFs ou ZIP com PDFs)
- **Relatório Excel** de processamento
- **OCR automático** para PDFs escaneados (Tesseract)

## Estrutura
- `app.py` — aplicação Streamlit
- `templates/MODELO_ATA.docx` — modelo de ata (Word)
- `requirements.txt` — dependências Python
- `packages.txt` — dependências Linux (Streamlit Cloud / Debian)
- `.streamlit/config.toml` — config do Streamlit

## Rodar localmente
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

> Para OCR/PDF em Linux local: instale também `tesseract-ocr`, `tesseract-ocr-por` e `poppler-utils`.
> Para gerar PDF a partir do DOCX: instale `libreoffice`.

## Deploy no Streamlit Community Cloud
1. Suba este repositório no GitHub
2. No Streamlit Cloud: **New app** → selecione repo → `app.py`
3. O Streamlit Cloud usa automaticamente:
   - `requirements.txt` (pip)
   - `packages.txt` (apt)

Pronto.
