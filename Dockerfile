# Lightweight Python runtime for the reporting application
FROM python:3.12-slim

WORKDIR /app

# Install dependencies first so Docker can cache this layer when
# application source changes without changing requirements.txt.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application source and supporting files.
COPY . .

# Flask/Gunicorn listens on port 5001 inside the container.
EXPOSE 5001

# Use Gunicorn for the containerized application rather than Flask's
# development server.
CMD ["gunicorn", "-w", "2", "-b", "0.0.0.0:5001", "app:app"]
