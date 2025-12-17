# Criador de Atas (Refatorado)

## O que faz
- Usa o template DOCX formatado em `templates/TEMPLATE_ATA.docx`.
- Extrai dados do contrato (PDF/DOCX). Para PDF escaneado, usa OCR automaticamente.
- Opcional: usa GPT (OpenAI) com saída estruturada (JSON Schema) para melhorar extração de sócios.
- Suporta lote (multi-upload ou ZIP) e gera relatório Excel.

## Rodar local
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy no Streamlit Community Cloud
- `requirements.txt`: dependências Python
- `packages.txt`: dependências Linux (LibreOffice + OCR)

## Habilitar GPT
No Streamlit Cloud: Settings → Secrets
```toml
OPENAI_API_KEY="..."
```
No app, marque “Usar GPT para melhorar extração”.
