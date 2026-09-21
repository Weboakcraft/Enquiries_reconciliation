# Container image for the Flask build of the Lead Reconciler.
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PORT=8000

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Run history and generated reports live here. Mount a volume on this path,
# or every restart starts the week-on-week trend again from nothing.
RUN mkdir -p data/runs data/exports \
 && adduser --disabled-password --gecos "" --uid 10001 oak \
 && chown -R oak:oak /app
USER oak
VOLUME ["/app/data"]

EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD python -c "import urllib.request,os;urllib.request.urlopen('http://127.0.0.1:'+os.environ.get('PORT','8000')+'/api/runs').read()"

CMD ["gunicorn", "-c", "gunicorn.conf.py", "app:app"]
