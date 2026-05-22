# SHOWLOG · DataHub

A web application that turns raw **brand-ads** TV ad-monitoring Excel exports
into the **Zee-format** per-show table, accumulates many files into a single
database, lets you view per channel, and exports the result.

This is the Python-backend successor to the old browser-only `datahub.html`.
It replaces `localStorage` with a real **SQLite database**, so data is shared
between everyone who uses the server — not trapped in one browser.

---

## What it does

- **Upload & accumulate** — drop one or many brand-ads `.xlsx` files. Each is
  transformed and merged into one combined dataset. Re-uploading a file with
  the same name **replaces** its rows (no double-counting).
- **Channel filtering + mapper** — 7 channel buttons (Zee, Star Plus, Colors,
  Set, Sab, Star Bharat, &TV). The mapper assigns each button to a raw
  `CHNLNAME` value; the mapping is saved in the database.
- **Date-grouped view** — rows sorted by date with a separator row showing the
  `WK Code` of each date group.
- **Export to Excel / CSV** — date-grouped, with the banded header, bold
  `WK Code`, and red heat-shaded `Mins Missing`. Export respects the active
  channel filter.
- Filters (week / day / search), sortable columns, pagination, and a
  "hide empty break/seg columns" toggle.

---

## Requirements

- **Python 3.9+** installed and on PATH.
  Get it from <https://www.python.org/downloads/> — during install on Windows,
  tick **"Add Python to PATH"**.

---

## How to run (Windows)

1. Put the whole `showlog_app` folder somewhere on the server.
2. **Double-click `run.bat`.**
   - First run installs the dependencies automatically.
   - It then starts the app and prints the address.
3. Open a browser:
   - on the server itself: `http://localhost:5000`
   - from another PC on the network: `http://SERVER-IP:5000`
     (find the IP by running `ipconfig` on the server).
4. To stop the app, close the black window.

## How to run (Mac / Linux, or manually)

```bash
cd showlog_app
pip install -r requirements.txt
python -m waitress --host=0.0.0.0 --port=5000 app:app
```

For quick local testing you can also use `python app.py` (Flask dev server).


---

## Signing in

The app requires a login. Two accounts are created automatically on first run:

| Username | Password |
|----------|----------|
| `admin` | `showlog@2026` |
| `user1` | `datahub@2026` |

**Change these passwords** after first run. Credentials are managed with:

```
python auth.py list                      # show all usernames
python auth.py add <username> <password> # add or change a user
python auth.py remove <username>         # delete a user
```

Credentials are stored as salted hashes in `users.json` (not plain text).
Delete `users.json` to reset back to the two default accounts.

This is lightweight protection for a small trusted team on an internal
network. It is not a hardened public-internet authentication system.

---

## The database

Everything lives in a single file: **`showlog.db`** (created on first run,
next to `app.py`).

- **Back up** = copy `showlog.db` somewhere safe.
- **Reset** = stop the app, delete `showlog.db`, restart. (Or use the
  "Clear All Data" button in the app.)

SQLite is fine for a small team doing mostly uploads and viewing. If the team
grows and many people upload at the exact same time, the database can be moved
to PostgreSQL — all database code is isolated in `database.py`, so only that
one file would change.

---

## File map

| File | Purpose |
|------|---------|
| `app.py` | Flask server, API routes, login |
| `auth.py` | User credentials (login) |
| `transform.py` | The brand-ads → Zee-format transform (grouping, breaks, segments, WK Code, Mins Missing) |
| `database.py` | SQLite schema and all queries |
| `excelio.py` | Reading uploaded `.xlsx`, building the export |
| `templates/index.html` | The single-page UI |
| `templates/login.html` | The login page |
| `static/app.js` | Frontend logic |
| `run.bat` | Windows launcher |
| `requirements.txt` | Python dependencies |
| `showlog.db` | The database (created on first run) |
| `users.json` | Login credentials (created on first run) |

---

## The transform, briefly

1. Read every ad spot from the brand-ads sheet (the header row is found by
   locating the row containing `CHNLNAME`/`PROGNAME`).
2. Group spots by `channel + date + programme + prog start` → one show row.
3. Per show, bucket spots by `BRK NO`; each break's Start = earliest ad start,
   End = latest ad end.
4. Segments = content stretches between breaks (a show with N breaks has N+1
   segments).
5. `WK Code = "Week " & (WEEKNUM(date,16) − 1)`; `Start Time` = prog start
   rounded to the nearest 30 minutes; `Mins Missing` = gap from the previous
   show's end, per channel per date.
