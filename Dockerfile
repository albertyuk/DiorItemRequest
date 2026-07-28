FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8080

# Single worker: runs are memory-heavy (read-only parse of a ~70 MB map) and
# the Fly machine has 1 GB. Two threads let a second request queue politely.
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "--workers", "1", "--threads", "2", \
     "--timeout", "300", "app:app"]
