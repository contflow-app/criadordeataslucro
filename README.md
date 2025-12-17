# Gerador de ATA (Template correto + GPT opcional)

## O que faz
- Usa o template DOCX em `templates/TEMPLATE_ATA.docx` (já formatado).
- Extrai dados do contrato (PDF/DOCX) via Regex.
- Opcional: usa GPT (OpenAI) com saída estruturada (JSON schema) para melhorar extração.
- Suporta lote (multi-upload ou ZIP) e gera relatório Excel.

## Rodar local
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy no Streamlit Community Cloud
- `requirements.txt` instala libs Python
- `packages.txt` instala LibreOffice + OCR (Tesseract + Poppler)

## Habilitar GPT
1. Crie `.streamlit/secrets.toml` com:
   ```toml
   OPENAI_API_KEY="..."
   ```
2. No app, marque "Usar GPT para melhorar extração".


## Padrões
- Cidade/UF: extraída da sede quando possível.
- Data da ATA: padrão = data do dia (Streamlit).
