# Gerador de ATA — PL 1087/2025 (Streamlit)

Este app gera uma ATA no formato do modelo de referência **"ATA DE SÓCIOS"** (estrutura similar ao PDF de exemplo).

## Rodar local
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy no Streamlit Community Cloud
- `requirements.txt`: dependências Python
- `packages.txt`: dependências Linux (OCR + LibreOffice para PDF)

## OCR
O app tenta extrair texto normal do PDF; se vier curto, aplica OCR (Tesseract).
