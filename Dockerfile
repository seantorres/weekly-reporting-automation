# --- Base image: a lightweight, official Python build ---
# "slim" trims unneeded OS packages to keep the image small.
FROM python:3.12-slim

# All following commands run from this directory inside the container
WORKDIR /app

# Copy just the dependency list first (not the whole app yet).
# Docker caches each step: as long as requirements.txt doesn't change,
# it reuses this layer on future builds instead of reinstalling everything.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Now copy the rest of the application code
COPY . .

# Documents which port the app listens on (informational — doesn't
# actually publish it; that happens with -p at `docker run` time)
EXPOSE 5001

# Production start command: gunicorn instead of Flask's dev server.
# -w 2        -> 2 worker processes handling requests
# -b 0.0.0.0:5001 -> listen on all interfaces inside the container, port 5001
# app:app     -> the Flask `app` object inside app.py
CMD ["gunicorn", "-w", "2", "-b", "0.0.0.0:5001", "app:app"]
