# ── Stage 1: build dependencies ─────────────────────────────────────────────
FROM python:3.12-slim AS builder

WORKDIR /app

# System deps for matplotlib (font rendering) and pandas
RUN apt-get update && apt-get install -y --no-install-recommends \
        libfreetype6 \
        fontconfig \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir --prefix=/install -r requirements.txt


# ── Stage 2: runtime image ───────────────────────────────────────────────────
FROM python:3.12-slim

WORKDIR /app

# Minimal runtime system libraries
RUN apt-get update && apt-get install -y --no-install-recommends \
        libfreetype6 \
        fontconfig \
    && rm -rf /var/lib/apt/lists/*

# Copy installed packages from builder
COPY --from=builder /install /usr/local

# Copy application source
COPY app.py data_loader.py generate_cbr.py chart_builder.py ./
COPY templates/ templates/
COPY static/ static/

# Copy the local template (used when TEMPLATE_BUCKET is not set)
# If you store the template in GCS, this layer can be removed.
COPY template/ template/

# Cloud Run injects PORT; gunicorn binds to it
ENV PORT=8080

# Use gunicorn for production; 2 workers is plenty for Cloud Run's single-container model.
# --timeout 300 gives up to 5 min for large decks.
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT} --workers 2 --timeout 300 app:app"]
