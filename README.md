# Weekly Reporting Automation

A Flask-based reporting application that automates a multi-source weekly reporting workflow. The application reconciles operational exports against a staff/team structure, applies source-specific cleaning and deduplication rules, builds reporting metrics, and generates an auditable Excel workbook.

> **Data privacy:** This repository is designed to use synthetic/sample data only. No real client records or confidential organizational data should be committed to the repository.

## What the application does

The application replaces a repetitive spreadsheet workflow with a browser-based reporting pipeline.

- Upload weekly CSV/XLSX source files through a Flask web interface.
- Maintain separate reporting structures for supported programs.
- Reconcile staff/service-user records against a team roster.
- Apply source-specific cleaning, filtering, and deduplication logic.
- Configure reporting metrics and date ranges.
- Generate a completed Excel workbook containing the weekly reporting table and supporting audit/pivot tabs.
- Save reusable program configuration and presets locally.
- Provide a synthetic-data preview for testing the workflow without production data.

## Architecture

```text
CSV / XLSX exports
        |
        v
Flask upload/API layer
        |
        +--> Structure / team reconciliation
        |
        +--> CLS processing
        |
        +--> Services processing
        |
        +--> Enrollment / outcomes processing
        |
        +--> Timeliness processing
        |
        v
pandas cleaning / deduplication / aggregation
        |
        v
Weekly reporting table + audited Excel workbook
```

The application separates processing responsibilities into service modules under `services/`. This keeps source-specific transformation rules isolated and makes the pipeline easier to maintain and test.

## Technology

- **Python 3.12**
- **Flask 3.1** — web application and API
- **pandas 2.3** — data cleaning, transformation, deduplication, and aggregation
- **openpyxl 3.1** — Excel workbook generation
- **Gunicorn** — production WSGI server inside the container
- **Docker / Docker Compose** — reproducible application environment
- **pytest** — automated tests
- **HTML / CSS / JavaScript** — browser interface

## Project structure

```text
.
├── app.py
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
├── .dockerignore
├── .gitignore
├── services/
│   ├── __init__.py
│   ├── cls_service.py
│   ├── enrollment_service.py
│   ├── services_service.py
│   ├── structure_service.py
│   └── timeliness_service.py
├── templates/
│   └── index.html
├── static/
│   ├── css/
│   │   └── styles.css
│   └── js/
│       └── app.js
├── data/
├── sample_data/
└── tests/
    └── test_app.py
```

## Run with Docker

Docker is the recommended way to run the application because it gives the application a consistent Python/runtime environment.

```bash
docker compose up --build
```

Then open:

```text
http://localhost:5001
```

Stop the application with:

```bash
docker compose down
```

The Docker image uses Gunicorn rather than Flask's development server:

```text
2 workers
0.0.0.0:5001
app:app
```

The Dockerfile installs dependencies before copying the application source so Docker can reuse the dependency layer when application code changes without changing `requirements.txt`.

## Run locally without Docker

Create and activate a virtual environment:

```bash
python -m venv .venv
```

macOS/Linux:

```bash
source .venv/bin/activate
```

Windows:

```powershell
.venv\Scripts\activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Start the application:

```bash
python app.py
```

Open:

```text
http://127.0.0.1:5001
```

## Tests

Run the automated test suite with:

```bash
pytest
```

The tests exercise Flask application behavior and API endpoints without requiring production data.

## Synthetic data

The `sample_data/` directory contains synthetic CSV/XLSX files used to exercise the reporting pipeline. These files are intentionally non-production examples and are included so the project can be demonstrated without exposing client information.

## Design decisions

### Auditability over prediction

The purpose of this application is reliable operational reporting rather than predictive modeling. The source systems contain changing definitions, historical data-quality issues, and organization-specific business rules. The application therefore prioritizes deterministic transformations, explicit rules, reproducible outputs, and traceability.

### Source-specific processing

Different exports have different schemas and business rules, so each major source has its own service module. This avoids putting all transformations into one large application function and makes individual processing paths easier to test and modify.

### Configuration instead of hard-coded reporting structure

Program structure, staff mappings, and reporting metrics are represented as configurable data where practical. This allows the application to support changes in team structure without rewriting the core reporting pipeline.

## Containerization notes

The included `Dockerfile` uses `python:3.12-slim` and installs the pinned Python dependencies from `requirements.txt`. The application listens on port `5001` inside the container.

For development, the container can be rebuilt with:

```bash
docker compose build --no-cache
docker compose up
```

## Disclaimer

This project is a generalized portfolio implementation inspired by real-world operational reporting workflows. All data used in this repository is synthetic or anonymized, and the project does not contain confidential, proprietary, or personally identifiable information.
