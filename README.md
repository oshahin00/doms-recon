# OpsFlow Studio

OpsFlow Studio is an enterprise-grade, containerized operations workflow and incident management platform built with Python, Streamlit, and PostgreSQL[cite: 1]. It streamlines complex hardware ticket reconciliations, MYF audits, and downstream incident cross-referencing for technical operations teams.

---

## ✨ Key Features

*   **DOMS Data Merge Tool:** Automated reconciliation and strict filtering for DOMS and Pump hardware tickets, systematically excluding irrelevant data (such as RFID or offline exceptions)[cite: 1].
*   **MYF Audit & Editor:** Interactive tabular editor allowing engineers to clean data, modify timestamps, and manage unique ticket tracking[cite: 1].
*   **Incident Cross-Matcher:** Cross-references raw incident reports against audited MYF datasets to isolate and format matched tickets[cite: 1].
*   **Stateful Run History:** Persists all extraction runs, record counts, and binary report snapshots directly to a PostgreSQL database, enabling easy review, deletion, or re-exporting[cite: 1].
*   **Automated Local Export:** Directly drops generated Excel spreadsheets into a synced local `output/` directory on the host machine without browser dialog interruptions[cite: 1, 4].
*   **Modern SaaS UI/UX:** Features a polished, high-contrast interface designed with custom CSS typography, card containers, and responsive layouts[cite: 1].

---

## 🛠️ Technology Stack & Requirements

*   **Frontend & Logic:** Python 3.9+, Streamlit[cite: 1, 2, 3]
*   **Data Processing:** Pandas, Openpyxl[cite: 1, 3]
*   **Database ORM:** SQLAlchemy, Psycopg2-binary (PostgreSQL Desktop)[cite: 1, 3]
*   **Containerization:** Docker Desktop (`python:3.9-slim`)[cite: 1, 2]
*   **Automation:** Windows Batch Scripting (`.bat`)

---

## 📂 Project Structure

```text
OpsFlow-Studio/
├── app.py              # Core Streamlit application & processing logic[cite: 1]
├── config.yaml         # Database credentials and environment configuration
├── Dockerfile          # Container build instructions[cite: 2]
├── requirements.txt    # Python package dependencies[cite: 3]
├── RunOpsFlow.bat      # Automated startup and container deployment script
├── StopOpsFlow.bat     # Teardown and container cleanup script
└── output/             # Synced local folder for generated Excel reports