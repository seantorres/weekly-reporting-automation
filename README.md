# Weekly Reporting App

A Flask application I built for my work as a Data Analyst at a homeless-services nonprofit, to replace a manual weekly reporting process that was taking up to half a day and prone to copy-paste errors.

The app pulls in four separate, messy source reports (case management exports, service logs, enrollment/outcomes data, and case-note timeliness data), reconciles them against a single staff/team roster, and produces the same operational summary table our program managers used to build by hand in a spreadsheet — plus a full audit-trail workbook so every number can be traced back to its source rows.

All data in this repo is synthetic, generated to mirror the shape and quirks of the real source files without containing any real client information.

## Why this exists

Two teams inside the organization each ran a version of the same weekly reporting process by hand: pulling exports from a case management system, cross-referencing them against a staff roster, and manually building a table of metrics per manager and per program. It was slow, easy to get wrong, and produced no audit trail if a number looked off.

This app automates that pipeline end-to-end: drop in the four source files, and it produces the finished report plus every intermediate calculation, in a fraction of the time.

## What it does

- **Reconciles four independent data sources** against one staff/team roster — case management service logs, program services data, enrollment/outcomes data, and case-note timeliness data — each with its own cleaning, filtering, and deduplication rules.
- **Handles messy real-world identity matching**: usernames that don't match full names, staff who moved teams mid-period, shared client records credited to multiple staff, and inactive staff who still show up in historical data.
- **Applies a "verified data wins" precedence rule**: automated exports are treated as a starting point, not ground truth, and known-correct manual corrections take priority over conflicting scraped/exported values.
- **Produces a complete audited output**: a finished weekly table (with manager subtotals and program totals) plus a separate workbook tab for every pivot and cleaning step, so any number can be traced back to its raw source rows.
- **Supports two independent program structures** (two different internal teams) with fully separate data, presets, and uploads, so a file from one program can never accidentally get applied to the other.

## Why there's no predictive/ML component

I considered adding a forecasting or predictive layer on top of this data and decided against it: the underlying source fields and definitions change depending on the reporting county, and the reliable historical window is only about five years with some known data-quality issues in older records. Rather than build a model on data that can't reliably support one, I kept this project focused on what the data quality actually allows — reliable, auditable reporting — and I'm exploring predictive work on a dataset better suited to it (see my fast-food price comparison project for that).

## Tech stack

- **Backend**: Python, Flask, served in production by gunicorn
- **Data processing**: pandas (multi-source cleaning, deduplication, pivoting)
- **Output**: openpyxl-generated Excel workbooks with formatted tables, subtotals, and audit tabs
- **Frontend**: vanilla HTML/CSS/JS
- **Containerization**: Docker / Docker Compose
- **Tests**: pytest

## Architecture

```
Source files (CSV/XLSX exports)
        |
Per-source cleaning & deduplication (pandas)
        |
Reconciliation against staff/team roster
        |
Pivot construction (per metric)
        |
Weekly summary table + audit workbook (Flask + openpyxl)
```

Each data source has its own service module (`services/`) with its own cleaning and matching rules, all reconciled against a single roster that acts as the source of truth for staff structure. Adding a new data source means adding a new service module — the roster, the table rendering, and the other sources don't need to change.

## Running it with Docker (recommended)

```bash
docker compose up --build
```

Open `http://localhost:5001`. Generated reports and saved presets persist in `outputs/` and `data/` on your machine between runs.

To stop it: `docker compose down`.

Under the hood, the container runs the app with gunicorn rather than Flask's built-in development server, since Flask's dev server isn't intended for anything beyond local development.

## Running it locally without Docker

```bash
cd weekly_reporting_app
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open `http://127.0.0.1:5001`.

On Windows, activate the environment with `.venv\Scripts\activate`.

Sample synthetic data is included in `sample_data/` to try the full pipeline end-to-end without any real files.

## Tests

```bash
pytest
```
