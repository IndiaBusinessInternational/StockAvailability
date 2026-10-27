# IBI Stock Availability v2.13

v2.13 (27 Oct 2026): edge to edge on every screen size: no centred 1180 px column on laptops; full-bleed boxes with 10 px inner padding; the add-stock form puts four fields per row on wide screens.

v2.12 (27 Oct 2026): top bar fits a 360 px phone with the Share button: the title shortens with "…", the Dark/Light word hides, the version badge always shows.

v2.11 (27 Oct 2026): new "Share this app" button in the top bar (ibi-share.js).

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
