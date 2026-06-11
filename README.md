# My Flower — Revenue System

A small web app for the My Flower hotel: upload a check-in export (`.xlsx`)
and get a clean booking revenue ledger.

## What it does

- **Cleans the file automatically** — repeated header rows, blank lines and the
  "Property Name / Email / Total Revenue" summary blocks are removed; only rows
  with a real booking number, check-in date and price are kept.
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
```

## Stack

- [React 18](https://react.dev) + [Vite](https://vitejs.dev)
- [SheetJS (xlsx)](https://www.npmjs.com/package/xlsx) for reading Excel exports
