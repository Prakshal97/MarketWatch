# NSE/BSE Announcement Tracker — Setup Guide

## What this does
- Fetches announcements from NSE + BSE public APIs
- Uses Claude AI to classify Auth Capital announcements vs others
- Auto-fills all fields: Board Approval, DOBM, CMP, MCap, Sector, Remarks, Action
- Generates a styled Excel with 2 sheets
- Web dashboard with filters + one-click download

---

## STEP 1 — Prerequisites

Install Python 3.10+ from https://python.org

---

## STEP 2 — Backend Setup

```bash
cd backend

# Install dependencies
pip install -r requirements.txt

# Create your .env file
cp .env.example .env
```

Now open `.env` and add your Anthropic API key:
```
ANTHROPIC_API_KEY=sk-ant-xxxxxxxxxxxxx
```

Get your API key from: https://console.anthropic.com

---

## STEP 3 — Run the Backend

```bash
cd backend
uvicorn main:app --reload --port 8000
```

You should see:
```
INFO:     Uvicorn running on http://127.0.0.1:8000
```

Test it: http://localhost:8000/health

---

## STEP 4 — Open the Dashboard

Just open `frontend/index.html` in your browser.
(Double-click the file, or drag it into Chrome/Firefox)

---

## STEP 5 — Use It

1. Select From Date and To Date
2. Click "⚡ Fetch Announcements"
3. Watch the progress bar
4. When done, see announcements in the table
5. Click "↓ Download Excel" to get the file

---

## Excel Output

The Excel file has 2 sheets:
- **Auth Capital** — all authorised capital increase announcements
- **Other Announcements** — everything else

Both sheets have the same columns:
Sr.no | Date of Entry | Name of the Company | Board Approval | DOBM |
Exst Auth Eq Cap (INR) | New Auth Eq Cap (INR) | Proposed Increase (INR) |
CMP | M Cap (in Cr) | Sector | Remark Positive | Remark Negative | Action | Link

---

## Note on API Key

If you don't set an Anthropic API key, the app still works!
It will use keyword-based classification (less accurate than AI).
Fields like Remarks and Action will be empty without the API key.

---

## Deploying to a Server (later)

When you're ready to host this online:

Backend: Deploy to any VPS (DigitalOcean, Linode, AWS EC2)
```bash
uvicorn main:app --host 0.0.0.0 --port 8000
```

Frontend: Update the `API` variable in index.html from:
```js
const API = 'http://localhost:8000';
```
to your server IP:
```js
const API = 'http://YOUR_SERVER_IP:8000';
```

Then just serve index.html from Nginx or any static host.

---

## Troubleshooting

**"Backend Offline" in dashboard**
→ Make sure uvicorn is running on port 8000

**NSE/BSE returning empty results**
→ NSE/BSE sometimes block automated requests. Try again after a few minutes.
→ NSE especially requires a valid browser session cookie.

**CMP not showing**
→ NSE/BSE quote APIs may be rate-limited. Data will show as "—" in that case.
