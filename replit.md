# AI.SEE Sales KPI Dashboard

## Overview
A web-based sales performance dashboard that visualizes KPIs across marketing funnels and revenue targets. Aggregates data from Google Sheets and HubSpot CRM.

## Tech Stack
- **Runtime**: Node.js 18+
- **Framework**: Express.js
- **Frontend**: Vanilla JavaScript, HTML5, CSS3
- **Charts**: Chart.js v4.4.0
- **APIs**: Google Sheets API, HubSpot CRM API
- **Config**: dotenv for environment variables

## Project Structure
```
server.js          # Main Express server, API endpoints
public/
  index.html       # Dashboard UI
  dashboard.js     # Client-side chart logic, tab switching
  styles.css       # Custom styles
package.json       # Dependencies
```

## Running the App
- Start: `node server.js`
- Runs on port 5000 (0.0.0.0)

## Environment Variables / Secrets Required
- `GOOGLE_CREDENTIALS` — JSON string of Google service account credentials (or use `credentials.json` file)
- `HUBSPOT_ACCESS_TOKEN` — HubSpot private app token
- `SPREADSHEET_ID` — (optional) Google Sheets ID, defaults to hardcoded value
- `SHEET_UEBERSICHT`, `SHEET_MARKETING`, `SHEET_STAGE2`, `SHEET_STAGE3`, `SHEET_REBUY` — (optional) Sheet tab names

## Dashboard Sections
- **Übersicht** — Overview of all sales stages
- **Marketing** — Marketing funnel metrics
- **Stage 2 (PoC)** — Proof of concept stage
- **Stage 3 (Follow-up)** — Follow-up stage
- **Rebuy Funnel** — Customer rebuy tracking

## Data Flow
1. Backend fetches data from Google Sheets via `googleapis`
2. Data is parsed/transformed by `parseSheet` helper
3. Frontend fetches `/api/data` and renders Chart.js visualizations
4. HubSpot deals fetched via `/api/hubspot/deals` for deal detail modals
