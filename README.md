# Retail Flow AI

Retail Flow AI is a garment retail inventory analysis app. Upload a purchase bill and a sales report to review item movement, remaining stock, profit, and suggested reorder actions. The app produces an Excel report and an interactive dashboard.

## Demo video

<video controls src="https://photos.app.goo.gl/JT5d2RURoJLQRqry5"></video>

## Features

- Enter retailer details to include in the generated report.
- Upload purchase and sales spreadsheets and view an analysis summary.
- Review purchase quantity, sold quantity, remaining stock, sales percentage, profit, and margin by item.
- Get rule-based value categories and reorder recommendations:
  - **High Value:** at least 70% sold and at least 25% margin; reorder immediately.
  - **Moderate Value:** at least 40% sold; order a limited quantity.
  - **Low Value:** below the moderate sales threshold; do not reorder.
- Download a formatted Excel report and open an interactive dashboard.

## Tech stack

- Frontend: React, Vite
- Backend: Python, Flask, pandas
- Reports: Excel (`openpyxl`) and interactive HTML dashboard (Plotly)

## Requirements

- Python 3
- Node.js and npm

## Run locally

Run the backend from the project root so generated files are saved in the expected `backend` folders.

### 1. Install backend dependencies

From the project root:

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r backend\requirements.txt
```

### 2. Start the backend

In a terminal at the project root:

```powershell
.\.venv\Scripts\python backend\app.py
```

The Flask API runs at `http://127.0.0.1:5000`.

### 3. Install and start the frontend

Open a second terminal:

```powershell
cd frontend
npm install
npm run dev
```

Open the local URL printed by Vite (normally `http://localhost:5173`).

## Spreadsheet format

Use the column names below in the uploaded Excel files. Column names are normalized for case and surrounding whitespace, but the required names must be present.

**Purchase bill**

`ITEM ARTICLE`, `QTY`, `SIZE`, `PURCHASE PRICE`, `COLOUR`, `TOTAL AMT`, `CATEGORY`, `FIT`, `MRP`

**Sales report**

`ITEM NAME`, `QTY`, `MRP`, `DISCOUNT AMOUNT`, `COLOUR`, `NET AMOUNT`, `CATEGORY`, `MARGIN`, `BILL NO`, `DATE`, `TAX AMOUNT`, `SIZE`

Sample spreadsheets are available in [`sampledata/`](./sampledata/).

## Generated files

After a successful analysis, the backend creates:

- `backend/reports/final_analysis_report.xlsx` — summary and item analysis worksheets.
- `backend/reports/stockpulse_dashboard.html` — interactive dashboard.

The app provides these through the **Download Excel** and **View Dashboard** actions.
