# My Flower — Revenue System

A small web app for the My Flower hotel: upload a check-in export (`.xlsx`)
and get a clean booking revenue ledger.

## Expected file

A `Checkin_…xlsx` export with these columns (header names are matched loosely):

`Book number · Guest name (Surname, Name) · Check-in · Check-out · Rooms ·
Price · Booker country · Unit type · Duration (nights)`

The file may contain several period blocks back to back — each with its own
repeated header row, blank separator rows and a "Property Name / Property
Number / Email / Dates / Total Revenue" summary block. Prices may be numbers
or text like `31$`; countries are 2-letter codes (`iq`, `ir`, `gb`, …);
multi-room bookings list their unit types together (`Quadruple Suite, Twin`).

## What it does

- **Cleans the file automatically** — repeated header rows, blank lines and the
  "Property Name / Email / Total Revenue" summary blocks are removed; only rows
  with a real booking number, check-in date and price are kept.
- **Removes duplicates** — the export's period blocks can overlap, so the same
  booking number may appear twice; only the first occurrence is counted.
- **Understands multi-room bookings** — the room-type filter matches each room
  type inside a combined booking, "Room-nights sold" counts nights × rooms, and
  the bookings table shows a ×2 / ×3 badge.
- **Shows countries properly** — booker country codes become names with flags
  (🇮🇶 Iraq instead of `IQ`).
- **Summary cards** — total income, bookings, nights sold, average per booking.
- **Monthly income ledger** — revenue per month with the best month highlighted.
- **Filters** — search by guest or booking number, filter by room type, country
  and month range, or pick N random guests to build a selection.
- **Bookings table** — every booking with check-in date, nights, room, country
  and price; tap a guest's name to add them to the selection.
- **Persistence** — the uploaded data is saved in the browser (localStorage),
  so it's still there next time you open the app. "Clear data" removes it.

Everything runs in the browser. No server, and the booking file never leaves
your machine.

## Run it

```bash
npm install
npm run dev      # local development at http://localhost:5173
npm run build    # production build in dist/
npm start        # serve the production build (uses $PORT, default 3000)
```

## Deploy on Railway

The repo is ready for [Railway](https://railway.com) as-is — `railway.json`
tells it to build with `npm run build` and serve the `dist/` folder with
`npm start` on Railway's `$PORT`.

1. In Railway: **New Project → Deploy from GitHub repo** and pick this repo.
2. Wait for the first build/deploy to go green.
3. Open the service → **Settings → Networking → Generate Domain** to get a
   public URL.

Every push to the deployed branch redeploys automatically.

Or with the [Railway CLI](https://docs.railway.com/guides/cli):

```bash
railway init
railway up
railway domain
```

## Stack

- [React 18](https://react.dev) + [Vite](https://vitejs.dev)
- [SheetJS (xlsx)](https://www.npmjs.com/package/xlsx) for reading Excel exports
- [serve](https://www.npmjs.com/package/serve) for hosting the built app
