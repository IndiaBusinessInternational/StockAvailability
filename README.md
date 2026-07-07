# IBI Stock Availability

Stock &amp; inventory availability tracker for **India Business International** (Kanyakumari).

**Live:** https://stock.indiabusinessinternational.online/

## Stack
- Single-file PWA (`index.html`) — installable, offline-capable via service worker (`sw.js`), dark/light themes, live clock, summary metrics, add/edit/delete, category & stock-status filters, search, sort, CSV export, and a category-wise stock report (print / PDF / CSV).
- Backend: Google Apps Script web app (`IBIStockAvailability_GAS.gs`) reading & writing an existing Google Sheet.

## Data API (Apps Script, GET-based)
`?action=ping | getAll | add | update | delete` → JSON `{ status:'ok', ... }`

The sheet keeps its existing layout (two header rows, data from row 3):

`S.No | Category | Product | HSN | GST | Image | Packed | Loose | Damage | Date of Updation | Star Rating | Keywords | Amazon rate 1-5 + Avg | Flipkart rate 1-5 + Avg`

A hidden `ID` column (col Y) is appended on the right for stable row identity — existing
rows are auto-assigned an ID the first time the app loads. Amazon/Flipkart averages are
computed by the backend and written as plain numbers (no `#DIV/0!`).

Sheet: https://docs.google.com/spreadsheets/d/1bp0OpZJKWB3-xEDKAgk5EErQ4JVg3uwmF2QAM35K6lA/edit

## Deploy
- **Frontend** — hosted on GitHub Pages from `main` (custom domain via `CNAME`). Push to `main` → auto-rebuilds.
- **Backend** — open the Sheet → Extensions → Apps Script, paste `IBIStockAvailability_GAS.gs`, then **Deploy → New deployment → Web app** (*Execute as: Me*, *Access: Anyone*). Copy the `/exec` URL and paste it into the app via **Menu → Backend Connection** (stored in the browser). On a later code change use **Manage deployments → Edit → New version** to keep the same `/exec` URL.
