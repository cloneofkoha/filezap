FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY form_filler_engine.py .
COPY api.py .
COPY master_data.md .

CMD uvicorn api:app --host 0.0.0.0 --port ${PORT:-8000}
