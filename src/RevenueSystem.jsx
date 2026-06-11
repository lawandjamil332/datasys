import { useState, useMemo, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

/* ------------------------------------------------------------------ */
/* Parsing helpers                                                     */
/* ------------------------------------------------------------------ */

const EXCEL_EPOCH_OFFSET = 25569; // days between 1899-12-30 and 1970-01-01

function toDate(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === "number" && v > 20000 && v < 80000) {
    return new Date(Math.round((v - EXCEL_EPOCH_OFFSET) * 86400 * 1000));
  }
  if (typeof v === "string") {
    const s = v.trim();
    // try d-m-yyyy or d/m/yyyy
    const m = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    const d = new Date(s);
    if (!isNaN(d)) return d;
  }
  return null;
}

function parsePrice(v) {
  if (v == null) return null;
  if (typeof v === "number") return isFinite(v) ? v : null;
  const cleaned = String(v).replace(/[^0-9.]/g, "");
  if (!cleaned) return null;
  const n = parseFloat(cleaned);
  return isFinite(n) ? n : null;
}

function cleanWorkbook(workbook) {
  const rows = [];
  const skipped = { junk: 0, badDate: 0, badPrice: 0 };
  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: null });
    for (const r of raw) {
      // find columns loosely (manual files sometimes rename headers)
      const keys = Object.keys(r);
      const get = (frag) => {
        const k = keys.find((k) => k.toLowerCase().includes(frag));
        return k ? r[k] : null;
      };
      const book = String(get("book") ?? "").trim();
      if (!/^\d{6,}$/.test(book)) {
        skipped.junk++;
        continue; // repeated headers, Property Name / Email / Total Revenue blocks, blanks
      }
      const checkIn = toDate(get("check-in") ?? get("check in"));
      const checkOut = toDate(get("check-out") ?? get("check out"));
      const price = parsePrice(get("price"));
      if (!checkIn) { skipped.badDate++; continue; }
      if (price == null) { skipped.badPrice++; continue; }

      let nights = parsePrice(get("duration"));
      if (!nights && checkOut) {
        nights = Math.max(1, Math.round((checkOut - checkIn) / 86400000));
      }
      rows.push({
        book,
        guest: String(get("guest") ?? "").trim() || "—",
        checkIn: checkIn.toISOString().slice(0, 10),
        checkOut: checkOut ? checkOut.toISOString().slice(0, 10) : null,
        rooms: parsePrice(get("rooms")) || 1,
        price,
        country: String(get("country") ?? "").trim().toUpperCase() || "—",
        unit: String(get("unit") ?? "").trim() || "—",
        nights: nights || 1,
      });
    }
  }
  rows.sort((a, b) => a.checkIn.localeCompare(b.checkIn));
  return { rows, skipped };
}

/* ------------------------------------------------------------------ */
/* Formatting helpers                                                  */
/* ------------------------------------------------------------------ */

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

const monthKey = (iso) => iso.slice(0, 7); // YYYY-MM
const monthLabel = (key) => {
  const [y, m] = key.split("-");
  return `${MONTHS[+m - 1]} ${y}`;
};
const fmtMoney = (n) =>
  "$" + n.toLocaleString("en-US", { maximumFractionDigits: 0 });
const fmtDate = (iso) => {
  if (!iso) return "—";
  const [y, m, d] = iso.split("-");
  return `${d} ${MONTHS[+m - 1]} ${y}`;
};

/* ------------------------------------------------------------------ */
/* Storage (persists data between sessions, with graceful fallback)    */
/* ------------------------------------------------------------------ */

const STORAGE_KEY = "myflower-bookings";

async function saveBookings(rows, fileName) {
  try {
    window.localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({ rows, fileName, savedAt: Date.now() })
    );
  } catch (e) {
    /* storage unavailable or full — app still works for this session */
  }
}

async function loadBookings() {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

async function clearBookings() {
  try {
    window.localStorage.removeItem(STORAGE_KEY);
  } catch (e) {}
}

/* ------------------------------------------------------------------ */
/* Main component                                                      */
/* ------------------------------------------------------------------ */

export default function RevenueSystem() {
  const [rows, setRows] = useState(null);
  const [fileName, setFileName] = useState("");
  const [skipped, setSkipped] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  // filters
  const [search, setSearch] = useState("");
  const [selectedGuests, setSelectedGuests] = useState([]); // empty = all
  const [unitFilter, setUnitFilter] = useState("All");
  const [countryFilter, setCountryFilter] = useState("All");
  const [fromMonth, setFromMonth] = useState("");
  const [toMonth, setToMonth] = useState("");
  const [randomCount, setRandomCount] = useState(10);
  const [showAllRows, setShowAllRows] = useState(false);

  const fileInputRef = useRef(null);

  useEffect(() => {
    (async () => {
      const saved = await loadBookings();
      if (saved && saved.rows && saved.rows.length) {
        setRows(saved.rows);
        setFileName(saved.fileName || "saved data");
      }
      setLoading(false);
    })();
  }, []);

  async function handleFile(file) {
    setError("");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const { rows: clean, skipped: sk } = cleanWorkbook(wb);
      if (!clean.length) {
        setError("No valid booking rows found in this file. Make sure it has Book number, Check-in and Price columns.");
        return;
      }
      setRows(clean);
      setSkipped(sk);
      setFileName(file.name);
      resetFilters();
      await saveBookings(clean, file.name);
    } catch (e) {
      setError("Couldn't read that file. Please upload an .xlsx export.");
    }
  }

  function resetFilters() {
    setSearch("");
    setSelectedGuests([]);
    setUnitFilter("All");
    setCountryFilter("All");
    setFromMonth("");
    setToMonth("");
    setShowAllRows(false);
  }

  /* ---------------- derived data ---------------- */

  const allGuests = useMemo(() => {
    if (!rows) return [];
    return [...new Set(rows.map((r) => r.guest))].sort();
  }, [rows]);

  const allUnits = useMemo(() => {
    if (!rows) return [];
    return [...new Set(rows.map((r) => r.unit))].sort();
  }, [rows]);

  const allCountries = useMemo(() => {
    if (!rows) return [];
    return [...new Set(rows.map((r) => r.country))].sort();
  }, [rows]);

  const allMonths = useMemo(() => {
    if (!rows) return [];
    return [...new Set(rows.map((r) => monthKey(r.checkIn)))].sort();
  }, [rows]);

  const filtered = useMemo(() => {
    if (!rows) return [];
    const q = search.trim().toLowerCase();
    const guestSet = selectedGuests.length ? new Set(selectedGuests) : null;
    return rows.filter((r) => {
      if (guestSet && !guestSet.has(r.guest)) return false;
      if (q && !r.guest.toLowerCase().includes(q) && !r.book.includes(q)) return false;
      if (unitFilter !== "All" && r.unit !== unitFilter) return false;
      if (countryFilter !== "All" && r.country !== countryFilter) return false;
      const mk = monthKey(r.checkIn);
      if (fromMonth && mk < fromMonth) return false;
      if (toMonth && mk > toMonth) return false;
      return true;
    });
  }, [rows, search, selectedGuests, unitFilter, countryFilter, fromMonth, toMonth]);

  const stats = useMemo(() => {
    const total = filtered.reduce((s, r) => s + r.price, 0);
    const nights = filtered.reduce((s, r) => s + (r.nights || 0), 0);
    const byMonth = {};
    for (const r of filtered) {
      const k = monthKey(r.checkIn);
      if (!byMonth[k]) byMonth[k] = { revenue: 0, bookings: 0, nights: 0 };
      byMonth[k].revenue += r.price;
      byMonth[k].bookings += 1;
      byMonth[k].nights += r.nights || 0;
    }
    const months = Object.entries(byMonth)
      .map(([k, v]) => ({ key: k, ...v }))
      .sort((a, b) => a.key.localeCompare(b.key));
    const peak = months.reduce((best, m) => (m.revenue > (best?.revenue ?? -1) ? m : best), null);
    return { total, nights, count: filtered.length, months, peak };
  }, [filtered]);

  function pickRandomGuests() {
    const pool = [...allGuests];
    const n = Math.max(1, Math.min(randomCount || 1, pool.length));
    for (let i = pool.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [pool[i], pool[j]] = [pool[j], pool[i]];
    }
    setSelectedGuests(pool.slice(0, n));
    setSearch("");
  }

  function toggleGuest(name) {
    setSelectedGuests((prev) =>
      prev.includes(name) ? prev.filter((g) => g !== name) : [...prev, name]
    );
  }

  const filtersActive =
    search || selectedGuests.length || unitFilter !== "All" ||
    countryFilter !== "All" || fromMonth || toMonth;

  const maxMonthRevenue = Math.max(1, ...stats.months.map((m) => m.revenue));
  const visibleRows = showAllRows ? filtered : filtered.slice(0, 50);

  /* ---------------- render ---------------- */

  return (
    <div className="app">
      <style>{css}</style>

      <header className="masthead">
        <div className="brand">
          <span className="bloom" aria-hidden="true">✿</span>
          <div>
            <h1>My Flower</h1>
            <p className="sub">Booking revenue ledger</p>
          </div>
        </div>
        {rows && (
          <div className="filemeta">
            <span className="filename" title={fileName}>{fileName}</span>
            <button className="ghost" onClick={() => fileInputRef.current?.click()}>
              Upload new file
            </button>
            <button
              className="ghost danger"
              onClick={async () => {
                await clearBookings();
                setRows(null);
                setFileName("");
                setSkipped(null);
                resetFilters();
              }}
            >
              Clear data
            </button>
          </div>
        )}
      </header>

      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls,.xlsm"
        style={{ display: "none" }}
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) handleFile(f);
          e.target.value = "";
        }}
      />

      {loading ? (
        <div className="empty"><p>Loading…</p></div>
      ) : !rows ? (
        <div
          className="empty drop"
          onDragOver={(e) => e.preventDefault()}
          onDrop={(e) => {
            e.preventDefault();
            const f = e.dataTransfer.files?.[0];
            if (f) handleFile(f);
          }}
        >
          <span className="bloom big" aria-hidden="true">✿</span>
          <h2>Upload your check-in export</h2>
          <p>
            Drop the .xlsx file here. Repeated headers, blank rows and the
            "Property / Email / Total Revenue" blocks are removed automatically —
            only real bookings are kept.
          </p>
          <button className="primary" onClick={() => fileInputRef.current?.click()}>
            Choose file
          </button>
          {error && <p className="error">{error}</p>}
        </div>
      ) : (
        <>
          {skipped && (
            <p className="cleannote">
              Loaded <strong>{rows.length.toLocaleString()}</strong> bookings ·
              removed {skipped.junk + skipped.badDate + skipped.badPrice} non-booking rows
            </p>
          )}
          {error && <p className="error">{error}</p>}

          {/* ----- summary cards ----- */}
          <section className="cards">
            <div className="card hero">
              <span className="label">{filtersActive ? "Income (selection)" : "Total income"}</span>
              <span className="value">{fmtMoney(stats.total)}</span>
            </div>
            <div className="card">
              <span className="label">Bookings</span>
              <span className="value">{stats.count.toLocaleString()}</span>
            </div>
            <div className="card">
              <span className="label">Nights sold</span>
              <span className="value">{stats.nights.toLocaleString()}</span>
            </div>
            <div className="card">
              <span className="label">Avg per booking</span>
              <span className="value">
                {stats.count ? fmtMoney(stats.total / stats.count) : "—"}
              </span>
            </div>
          </section>

          {/* ----- filters ----- */}
          <section className="panel">
            <div className="panelhead">
              <h3>Build a selection</h3>
              {filtersActive && (
                <button className="ghost" onClick={resetFilters}>Reset — show everything</button>
              )}
            </div>
            <div className="filters">
              <label className="field grow">
                <span>Search guest or booking no.</span>
                <input
                  type="text"
                  value={search}
                  placeholder="e.g. Alnajjar"
                  onChange={(e) => setSearch(e.target.value)}
                />
              </label>
              <label className="field">
                <span>Room type</span>
                <select value={unitFilter} onChange={(e) => setUnitFilter(e.target.value)}>
                  <option>All</option>
                  {allUnits.map((u) => <option key={u}>{u}</option>)}
                </select>
              </label>
              <label className="field">
                <span>Country</span>
                <select value={countryFilter} onChange={(e) => setCountryFilter(e.target.value)}>
                  <option>All</option>
                  {allCountries.map((c) => <option key={c}>{c}</option>)}
                </select>
              </label>
              <label className="field">
                <span>From month</span>
                <select value={fromMonth} onChange={(e) => setFromMonth(e.target.value)}>
                  <option value="">Start</option>
                  {allMonths.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}
                </select>
              </label>
              <label className="field">
                <span>To month</span>
                <select value={toMonth} onChange={(e) => setToMonth(e.target.value)}>
                  <option value="">End</option>
                  {allMonths.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}
                </select>
              </label>
              <div className="field">
                <span>Random guests</span>
                <div className="randomrow">
                  <input
                    type="number"
                    min="1"
                    max={allGuests.length}
                    value={randomCount}
                    onChange={(e) => setRandomCount(parseInt(e.target.value) || 1)}
                  />
                  <button className="primary slim" onClick={pickRandomGuests}>Pick</button>
                </div>
              </div>
            </div>

            {selectedGuests.length > 0 && (
              <div className="chips">
                <span className="chiplabel">{selectedGuests.length} guest{selectedGuests.length > 1 ? "s" : ""} selected:</span>
                {selectedGuests.map((g) => (
                  <button key={g} className="chip" onClick={() => toggleGuest(g)} title="Remove">
                    {g} ✕
                  </button>
                ))}
                <button className="ghost slim" onClick={() => setSelectedGuests([])}>Clear selection</button>
              </div>
            )}
          </section>

          {/* ----- monthly ledger ----- */}
          <section className="panel">
            <div className="panelhead">
              <h3>Monthly income</h3>
              <span className="muted">{stats.months.length} month{stats.months.length !== 1 ? "s" : ""}</span>
            </div>
            {stats.months.length === 0 ? (
              <p className="muted">Nothing matches this selection.</p>
            ) : (
              <div className="ledger">
                {stats.months.map((m) => (
                  <div className="ledgerrow" key={m.key}>
                    <span className="lmonth">
                      {monthLabel(m.key)}
                      {stats.peak?.key === m.key && <span className="peak" title="Best month"> ✿</span>}
                    </span>
                    <div className="lbarwrap">
                      <div
                        className={"lbar" + (stats.peak?.key === m.key ? " best" : "")}
                        style={{ width: `${(m.revenue / maxMonthRevenue) * 100}%` }}
                      />
                    </div>
                    <span className="lrev">{fmtMoney(m.revenue)}</span>
                    <span className="lmeta">{m.bookings} bkg · {m.nights} nts</span>
                  </div>
                ))}
              </div>
            )}
          </section>

          {/* ----- bookings table ----- */}
          <section className="panel">
            <div className="panelhead">
              <h3>Bookings</h3>
              <span className="muted">
                showing {visibleRows.length.toLocaleString()} of {filtered.length.toLocaleString()}
              </span>
            </div>
            <div className="tablewrap">
              <table>
                <thead>
                  <tr>
                    <th>Guest</th>
                    <th>Check-in</th>
                    <th className="num">Nights</th>
                    <th>Room</th>
                    <th>Country</th>
                    <th className="num">Price</th>
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.map((r, i) => (
                    <tr key={r.book + i}>
                      <td>
                        <button
                          className={"linky" + (selectedGuests.includes(r.guest) ? " on" : "")}
                          onClick={() => toggleGuest(r.guest)}
                          title="Add or remove this guest from the selection"
                        >
                          {r.guest}
                        </button>
                      </td>
                      <td>{fmtDate(r.checkIn)}</td>
                      <td className="num">{r.nights}</td>
                      <td>{r.unit}</td>
                      <td>{r.country}</td>
                      <td className="num strong">{fmtMoney(r.price)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {filtered.length > 50 && (
              <button className="ghost wide" onClick={() => setShowAllRows((s) => !s)}>
                {showAllRows ? "Show first 50 only" : `Show all ${filtered.length.toLocaleString()} rows`}
              </button>
            )}
            <p className="hint">Tip: tap a guest's name to add them to the selection.</p>
          </section>
        </>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Styles                                                              */
/* ------------------------------------------------------------------ */

const css = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,640&family=Public+Sans:wght@400;500;600&display=swap');

:root {
  --paper: #F7F3EA;
  --ink: #22302B;
  --green: #2E4A3D;
  --rose: #C2566B;
  --brass: #B98A45;
  --line: #E2DAC8;
  --card: #FFFEF9;
}
* { box-sizing: border-box; }
body { margin: 0; }
.app {
  min-height: 100vh;
  background: var(--paper);
  color: var(--ink);
  font-family: 'Public Sans', system-ui, sans-serif;
  padding: 20px clamp(14px, 4vw, 44px) 60px;
  font-size: 15px;
}
h1, h2, h3 { font-family: 'Fraunces', serif; font-weight: 640; margin: 0; }

.masthead {
  display: flex; align-items: flex-end; justify-content: space-between;
  gap: 16px; flex-wrap: wrap;
  padding-bottom: 16px; margin-bottom: 20px;
  border-bottom: 2px solid var(--ink);
}
.brand { display: flex; align-items: center; gap: 12px; }
.brand h1 { font-size: clamp(26px, 5vw, 34px); line-height: 1; }
.sub { margin: 4px 0 0; color: var(--green); font-size: 13px; letter-spacing: .04em; text-transform: uppercase; }
.bloom { color: var(--rose); font-size: 28px; line-height: 1; }
.bloom.big { font-size: 52px; display: block; margin-bottom: 10px; }
.filemeta { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
.filename { font-size: 12px; color: var(--green); max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }

button { font: inherit; cursor: pointer; border-radius: 8px; }
.primary {
  background: var(--green); color: var(--paper); border: none;
  padding: 10px 20px; font-weight: 600;
}
.primary:hover { background: #243B31; }
.primary.slim { padding: 8px 14px; }
.ghost {
  background: transparent; border: 1px solid var(--line);
  color: var(--ink); padding: 7px 12px; font-size: 13px;
}
.ghost:hover { border-color: var(--green); }
.ghost.danger:hover { border-color: var(--rose); color: var(--rose); }
.ghost.slim { padding: 4px 10px; }
.ghost.wide { width: 100%; margin-top: 10px; }
button:focus-visible, input:focus-visible, select:focus-visible {
  outline: 2px solid var(--brass); outline-offset: 2px;
}

.empty {
  max-width: 460px; margin: 12vh auto 0; text-align: center;
}
.empty.drop {
  background: var(--card); border: 1.5px dashed var(--brass);
  border-radius: 16px; padding: 44px 28px;
}
.empty h2 { font-size: 24px; margin-bottom: 8px; }
.empty p { color: #5A6660; line-height: 1.5; margin: 0 0 20px; }
.error { color: var(--rose); font-weight: 500; margin-top: 14px; }
.cleannote { font-size: 13px; color: var(--green); margin: 0 0 14px; }

.cards {
  display: grid; gap: 12px;
  grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
  margin-bottom: 18px;
}
.card {
  background: var(--card); border: 1px solid var(--line);
  border-radius: 12px; padding: 14px 16px;
  display: flex; flex-direction: column; gap: 6px;
}
.card.hero { background: var(--green); border-color: var(--green); }
.card.hero .label { color: #CFE0D6; }
.card.hero .value { color: var(--paper); }
.label { font-size: 11px; text-transform: uppercase; letter-spacing: .08em; color: #6B756F; }
.value {
  font-family: 'Fraunces', serif; font-weight: 640;
  font-size: clamp(22px, 4vw, 30px); font-variant-numeric: tabular-nums;
}

.panel {
  background: var(--card); border: 1px solid var(--line);
  border-radius: 12px; padding: 16px; margin-bottom: 18px;
}
.panelhead {
  display: flex; justify-content: space-between; align-items: baseline;
  gap: 10px; margin-bottom: 14px; flex-wrap: wrap;
}
.panelhead h3 { font-size: 18px; }
.muted { color: #6B756F; font-size: 13px; }

.filters { display: flex; gap: 12px; flex-wrap: wrap; }
.field { display: flex; flex-direction: column; gap: 5px; min-width: 130px; }
.field.grow { flex: 1 1 200px; }
.field > span { font-size: 11px; text-transform: uppercase; letter-spacing: .07em; color: #6B756F; }
input[type="text"], input[type="number"], select {
  font: inherit; padding: 9px 11px; border: 1px solid var(--line);
  border-radius: 8px; background: #fff; color: var(--ink); width: 100%;
}
.randomrow { display: flex; gap: 8px; }
.randomrow input { width: 72px; }

.chips { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; margin-top: 14px; }
.chiplabel { font-size: 13px; color: var(--green); font-weight: 600; }
.chip {
  background: #F3E7D9; border: 1px solid var(--brass); color: var(--ink);
  padding: 4px 10px; border-radius: 999px; font-size: 13px;
}
.chip:hover { background: var(--rose); border-color: var(--rose); color: #fff; }

.ledger { display: flex; flex-direction: column; gap: 7px; }
.ledgerrow {
  display: grid; align-items: center; gap: 10px;
  grid-template-columns: 84px 1fr 76px 96px;
}
.lmonth { font-size: 13px; font-weight: 600; white-space: nowrap; }
.peak { color: var(--rose); }
.lbarwrap { background: #EFEADD; border-radius: 4px; height: 18px; overflow: hidden; }
.lbar {
  height: 100%; background: var(--brass); border-radius: 4px;
  min-width: 2px; transition: width .4s ease;
}
.lbar.best { background: var(--rose); }
.lrev { font-variant-numeric: tabular-nums; font-weight: 600; text-align: right; font-size: 14px; }
.lmeta { font-size: 11.5px; color: #6B756F; text-align: right; white-space: nowrap; }
@media (max-width: 560px) {
  .ledgerrow { grid-template-columns: 70px 1fr 70px; }
  .lmeta { display: none; }
}
@media (prefers-reduced-motion: reduce) { .lbar { transition: none; } }

.tablewrap { overflow-x: auto; }
table { width: 100%; border-collapse: collapse; font-size: 14px; }
th {
  text-align: left; font-size: 11px; text-transform: uppercase;
  letter-spacing: .07em; color: #6B756F; font-weight: 600;
  padding: 6px 10px; border-bottom: 1.5px solid var(--ink);
}
td { padding: 8px 10px; border-bottom: 1px solid var(--line); white-space: nowrap; }
.num { text-align: right; font-variant-numeric: tabular-nums; }
.strong { font-weight: 600; }
.linky {
  background: none; border: none; padding: 0; color: var(--ink);
  text-decoration: underline; text-decoration-color: var(--line);
  text-underline-offset: 3px;
}
.linky:hover { text-decoration-color: var(--rose); color: var(--rose); }
.linky.on { color: var(--rose); font-weight: 600; text-decoration-color: var(--rose); }
.hint { font-size: 12px; color: #6B756F; margin: 10px 0 0; }
`;
