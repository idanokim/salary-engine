# Salary Engine API — Pay Simulator

Multi-track salary simulator for Israeli civil-service pay (מנהלת הגמלאות).
Given a worker's track (דירוג), grade (דרגה), seniority (ותק) and job %, it
computes the expected salary from the official pay tables and flags each pay
slip as **valid (תקין)** or **invalid (שגוי)**. Built with FastAPI.

📖 **How the engine works:** see [`ENGINE.md`](ENGINE.md) — "the brain of the simulator".

The pay tables (75 grades, 9 seniority tracks, per-track caps) are extracted
from the Progim workbook into `lookups.json` via `tools/extract_lookups.py`,
and the component rules (החוקה — percentage bases/rates per pay code, plus the
manual 'ידני' codes) into `component_rules.json` via `tools/extract_rules.py`.

## Live endpoints (once deployed)

| Endpoint | Method | Description |
|---|---|---|
| `/` | GET | Web UI (single-worker calculator + file validation) |
| `/healthz` | GET | Health check |
| `/api/info` | GET | Engine status: grades, tracks, seniority caps |
| `/api/calculate` | POST | Calculate + validate one worker's salary |
| `/api/accuracy` | POST | Upload a גולמי .xlsx → valid/invalid slip stats |
| `/api/batch` | POST | Upload a גולמי .xlsx → results .xlsx with two tabs: **תקין** (valid) and **לבדיקה** (everything else) |
| `/api/export-highlighted` | POST | Upload a גולמי .xlsx → the same גולמי-מעודכן pivot back, with every invalid pay code and wrong total marked **yellow** |
| `/api/lookups` | GET | Full lookup tables (darga/vetek/caps) — lets the web UI run the engine **in the browser** |
| `/api/rules` | GET | Component rules (החוקה) from the Progim workbook — percentage bases/rates + manual (ידני) codes |

> **Browser-side file checking:** the web UI parses and validates גולמי files
> **locally in the browser** (SheetJS + `engine.js`, the same logic as the API),
> so large files are never uploaded — this sidesteps serverless request-body
> limits (e.g. Vercel rejects bodies over ~4.5 MB). The `/api/accuracy`,
> `/api/batch` and `/api/export-highlighted` endpoints remain for API clients.
| `/api/grades` | GET | All grade labels and base salaries |
| `/api/tracks` | GET | All seniority tracks and their ותק caps |
| `/api/vatek/{years}?track=N` | GET | Seniority multiplier for a track |
| `/docs` | GET | Interactive Swagger UI (try it in browser) |

---

## Deploy to Vercel (free)

The repo is Vercel-ready: `api/index.py` exposes the FastAPI app as a serverless
function, `vercel.json` rewrites every route to it (so the frontend at `/`, the
API under `/api/...`, and `/docs` all work), and the reference data
(`golmi.xlsx`) is bundled via `includeFiles`.

### Option A — Dashboard

1. Push this repo to GitHub.
2. Go to [vercel.com](https://vercel.com) → **Add New → Project** → import the repo.
3. Leave all build settings at their defaults (Vercel auto-detects the Python
   function and `requirements.txt`) and click **Deploy**.

Your app will be live at `https://<project>.vercel.app`:
- `/` — the interactive frontend (calls the API on the same origin)
- `/docs` — Swagger UI
- `/api/info`, `/api/calculate`, `/api/accuracy`, `/api/batch`, ...

### Option B — CLI

```bash
npm i -g vercel
vercel        # preview deploy
vercel --prod # production deploy
```

> Note: the batch/accuracy endpoints parse the full Excel file in-memory. On
> Vercel's Hobby plan the function timeout caps at 60s (`maxDuration` in
> `vercel.json`); very large uploads may need a paid plan.

---

## Deploy to Render (free, ~3 minutes)

### Step 1 — Push to GitHub

```bash
# In this folder:
git init
git add .
git commit -m "initial salary engine API"

# Create a new repo on github.com, then:
git remote add origin https://github.com/YOUR_USERNAME/salary-engine.git
git push -u origin main
```

### Step 2 — Connect to Render

1. Go to [render.com](https://render.com) and sign up / log in with GitHub
2. Click **New → Web Service**
3. Select your `salary-engine` repository
4. Render auto-detects `render.yaml` — just click **Deploy**
5. Wait ~2 minutes for the build

Your API will be live at:
```
https://salary-engine.onrender.com
```

### Step 3 — Test it

Open `https://salary-engine.onrender.com/docs` in your browser.
You'll see the full interactive Swagger UI — try the `/api/calculate` endpoint directly.

Or with curl:
```bash
curl https://salary-engine.onrender.com/api/info
```

---

## Run locally (optional)

```bash
pip install -r requirements.txt
uvicorn main:app --reload
# Open http://localhost:8000/docs
```

---

## Project structure

```
salary_api/
├── main.py              # FastAPI app — all endpoints
├── requirements.txt     # Python dependencies
├── render.yaml          # Render deployment config
├── engine/
│   ├── lookups.py       # Grade (דרגה) and seniority (ותק) table loader
│   ├── calculator.py    # Single-worker salary calculation
│   └── batch.py         # Batch runner over גולמי Excel files
└── data/
    └── golmi.xlsx       # Reference data (grade + seniority lookup tables)
```

## API usage example

**Calculate one worker:**

```bash
curl -X POST https://salary-engine.onrender.com/api/calculate \
  -H "Content-Type: application/json" \
  -d '{
    "worker_id": 11021106,
    "ministry_code": 170,
    "ministry_name": "מכס ומע\"מ",
    "droog": 1,
    "job_pct": 1.0,
    "kod_darga": 202,
    "darga_label": "18",
    "vatek": 33.75,
    "components": [
      {"code": 10002, "name": "שכר משולב", "amount": 4158.5, "pensionable": "כן"},
      {"code": 4544,  "name": "3.6% ת. שכר", "amount": 228.77, "pensionable": "כן"}
    ]
  }'
```

**Batch verify a file:**

```bash
curl -X POST https://salary-engine.onrender.com/api/batch \
  -F "file=@path/to/golmi.xlsx" \
  -o results.csv
```
