FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8080

# Single worker: runs are memory-heavy (read-only parse of a ~70 MB map) and
# the Fly machine has 1 GB. Two threads let a second request queue politely.
#
# Deliberately NOT set here:
#   --max-requests    recycling the worker would kill in-flight background
#                     scans (the job registry is in-process).
#   --access-logfile  request URLs carry single-use invite/reset tokens;
#                     logging them would persist working links to disk.
# gunicorn's default request-line/header limits (4094 B / 100 fields /
# 8190 B per field) are already the restrictive values, so they stand.
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "--workers", "1", "--threads", "2", \
     "--timeout", "300", "app:app"]
