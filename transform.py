"""
transform.py — Brand-ads INPUT -> Zee-format OUTPUT transform core.

Ported from the SHOWLOG / DataHub spec (handoff section 3).

INPUT  ("brand-ads"): one row = one ad spot. Header row is located by finding
       the row that contains CHNLNAME / PROGNAME (there are title rows above).
OUTPUT ("Zee TV"):    one row = one show, 68 columns, with up to 18 ad-break
       Start/End pairs and 19 segment durations.
"""

from datetime import datetime, timedelta, date
import math

# ---------------------------------------------------------------------------
# Output column layout (68 columns) — matches Output_Final.xlsx exactly.
#
# Show info (13): WK CODE, Mins Missing, Week, Day, Date, Start Time,
#                 PROGNAME, PROG ST, PROG ET, Content Dur, Prog Dur,
#                 Ad Dur, No.Of Breaks
# Ad breaks (36): 18 x (Ad ST, Ad ET)
# Segments  (19): Seg 1 .. Seg 18, Last Seg
# ---------------------------------------------------------------------------

SHOW_INFO_COLS = [
    "WK CODE", "Mins Missing", "Week", "Day", "Date", "Start Time",
    "PROGNAME", "PROG ST", "PROG ET", "Content Dur", "Prog Dur",
    "Ad Dur", "No.Of Breaks",
]

# Break columns: 18 pairs of "Ad ST" / "Ad ET". Internally we need unique
# keys, so each carries a 1-based break index suffix; the *display* label
# (in Excel and the UI table header) is just "Ad ST" / "Ad ET", with a
# "Break N" group header above.
BREAK_COLS = []
BREAK_COL_LABELS = {}   # internal key -> display label
BREAK_COL_GROUP = {}    # internal key -> "Break N"
for i in range(1, 19):
    st, et = f"Ad ST {i}", f"Ad ET {i}"
    BREAK_COLS.extend([st, et])
    BREAK_COL_LABELS[st] = "Ad ST"
    BREAK_COL_LABELS[et] = "Ad ET"
    BREAK_COL_GROUP[st] = f"Break {i}"
    BREAK_COL_GROUP[et] = f"Break {i}"

SEG_COLS = [f"Seg {i}" for i in range(1, 19)] + ["Last Seg"]

OUTPUT_COLS = SHOW_INFO_COLS + BREAK_COLS + SEG_COLS  # 13 + 36 + 19 = 68

# Display label for any column (break cols collapse to "Ad ST"/"Ad ET").
def display_label(col):
    return BREAK_COL_LABELS.get(col, col)

# "Channel" is no longer an output column, but it is still needed internally
# for channel filtering and the mapper. It is carried on each row under the
# private "_channel" key (see transform()).

# Column group boundaries for the banded header in the UI
COL_GROUPS = {
    "Show Info": SHOW_INFO_COLS,
    "Ad Breaks": BREAK_COLS,
    "Segments": SEG_COLS,
}

# Required input columns we look for to detect the header row
INPUT_HEADER_MARKERS = ["CHNLNAME", "PROGNAME"]

DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday",
             "Friday", "Saturday", "Sunday"]


# ---------------------------------------------------------------------------
# Time / date helpers
# ---------------------------------------------------------------------------

def _excel_serial_to_datetime(serial):
    """Convert an Excel date serial number to a datetime."""
    # Excel epoch is 1899-12-30 (accounts for the 1900 leap-year bug).
    return datetime(1899, 12, 30) + timedelta(days=float(serial))


def parse_date(value):
    """Parse a date cell into a datetime.date. Returns None on failure."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        try:
            return _excel_serial_to_datetime(value).date()
        except Exception:
            return None
    s = str(value).strip()
    if not s:
        return None
    fmts = ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y",
            "%d/%m/%y", "%d-%b-%Y", "%d-%b-%y", "%d.%m.%Y"]
    for fmt in fmts:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def parse_time_to_seconds(value):
    """Parse a time cell into seconds-since-midnight. Returns None on failure."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.hour * 3600 + value.minute * 60 + value.second
    if hasattr(value, "hour") and hasattr(value, "minute"):  # datetime.time
        return value.hour * 3600 + value.minute * 60 + value.second
    if isinstance(value, (int, float)):
        # Excel time = fraction of a day. A whole number could be seconds.
        if 0 <= value < 1:
            return round(value * 86400)
        if value >= 1:
            # Treat as seconds if plausible, else fraction-of-day remainder.
            if value < 86400:
                return int(round(value)) % 86400
            return round((value - math.floor(value)) * 86400)
        return None
    s = str(value).strip()
    if not s:
        return None
    parts = s.split(":")
    try:
        parts = [int(float(p)) for p in parts]
    except ValueError:
        return None
    if len(parts) == 3:
        h, m, sec = parts
    elif len(parts) == 2:
        h, m, sec = parts[0], parts[1], 0
    elif len(parts) == 1:
        h, m, sec = parts[0], 0, 0
    else:
        return None
    return (h * 3600 + m * 60 + sec) % 86400


def seconds_to_hms(seconds):
    """Format seconds-since-midnight as HH:MM:SS."""
    if seconds is None:
        return ""
    seconds = int(round(seconds)) % 86400
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


def parse_duration_to_seconds(value):
    """
    Parse a duration cell into SECONDS.

    A duration may arrive as:
      * an Excel time/duration serial (fraction of a day, 0..1)  -> * 86400
      * a time string "HH:MM:SS" or "MM:SS"                      -> parsed
      * a plain number -> AMBIGUOUS (could be seconds or minutes).
        It is returned AS-IS here, treated as seconds; the caller
        (transform) decides the unit by sanity-checking against the
        programme duration. See _resolve_units().
    """
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        if 0 < value < 1:                      # Excel fraction-of-day
            return float(value) * 86400.0
        return float(value)                    # plain number -> seconds (tentative)
    s = str(value).strip()
    if not s:
        return 0.0
    if ":" in s:
        return float(parse_time_to_seconds(s) or 0)
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_duration_to_minutes(value):
    """Backwards-compatible helper: parse a duration cell into minutes."""
    return parse_duration_to_seconds(value) / 60.0


def weeknum_mode16(d):
    """
    Excel WEEKNUM(date, 16): weeks start on Saturday.
    Week 1 is the week containing January 1.
    """
    jan1 = date(d.year, 1, 1)
    # Python weekday(): Mon=0..Sun=6. Saturday=5.
    # Days from the Saturday on/before Jan 1 to Jan 1.
    jan1_offset = (jan1.weekday() - 5) % 7
    week_start = jan1 - timedelta(days=jan1_offset)
    delta_days = (d - week_start).days
    return delta_days // 7 + 1


def wk_code(d):
    """WK Code = 'Week ' & (WEEKNUM(date,16) - 1)."""
    if d is None:
        return ""
    return f"Week {weeknum_mode16(d) - 1}"


def mround(value, multiple):
    """Excel MROUND: round to nearest multiple."""
    if multiple == 0:
        return 0
    return round(value / multiple) * multiple


def start_time_rounded(seconds):
    """Start Time = programme start rounded to nearest 30 minutes."""
    if seconds is None:
        return None
    return mround(seconds, 1800)  # 1800s = 30 min


# ---------------------------------------------------------------------------
# Reading the brand-ads workbook
# ---------------------------------------------------------------------------

def _norm(s):
    return str(s).strip().upper() if s is not None else ""


def find_header_row(rows):
    """
    Locate the header row: the first row containing the input header markers.
    `rows` is a list of lists (cell values). Returns (header_index, col_map)
    where col_map maps normalised column name -> column index.
    """
    for idx, row in enumerate(rows):
        norm_cells = [_norm(c) for c in row]
        if all(any(marker in cell for cell in norm_cells)
               for marker in INPUT_HEADER_MARKERS):
            col_map = {}
            for ci, cell in enumerate(norm_cells):
                if cell:
                    col_map[cell] = ci
            return idx, col_map
    return None, None


def _col(col_map, *names):
    """Find a column index by trying several candidate names."""
    for name in names:
        key = _norm(name)
        if key in col_map:
            return col_map[key]
        for k, v in col_map.items():
            if key in k:
                return v
    return None


def read_brand_ads(rows):
    """
    Parse brand-ads rows (list of lists) into a flat list of ad-spot dicts.
    Raises ValueError if the header row cannot be found.
    """
    header_idx, col_map = find_header_row(rows)
    if header_idx is None:
        raise ValueError(
            "Could not find the header row. Expected a row containing "
            "CHNLNAME and PROGNAME."
        )

    c = {
        "chnl":  _col(col_map, "CHNLNAME", "CHANNEL"),
        "pdate": _col(col_map, "PROG DATE", "PROGDATE"),
        "pst":   _col(col_map, "PROG ST", "PROGST"),
        "pet":   _col(col_map, "PROG ET", "PROGET"),
        "pname": _col(col_map, "PROGNAME", "PROG NAME"),
        "genre": _col(col_map, "PROG GENRE", "GENRE"),
        "pdur":  _col(col_map, "PROG DUR", "PROGDUR"),
        "brk":   _col(col_map, "BRK NO", "BREAK NO", "BRKNO"),
        "dur":   _col(col_map, "DURATION", "AD DURATION"),
        "adst":  _col(col_map, "ADST", "AD ST"),
        "adet":  _col(col_map, "ADET", "AD ET"),
        "brand": _col(col_map, "BRAND NAME", "BRANDNAME"),
    }

    spots = []
    for row in rows[header_idx + 1:]:
        if not row or all(cell in (None, "") for cell in row):
            continue
        chnl = row[c["chnl"]] if c["chnl"] is not None and c["chnl"] < len(row) else None
        pname = row[c["pname"]] if c["pname"] is not None and c["pname"] < len(row) else None
        if not chnl and not pname:
            continue

        def cell(key):
            ci = c[key]
            if ci is None or ci >= len(row):
                return None
            return row[ci]

        spot = {
            "channel": str(chnl).strip() if chnl else "",
            "date": parse_date(cell("pdate")),
            "prog_st": parse_time_to_seconds(cell("pst")),
            "prog_et": parse_time_to_seconds(cell("pet")),
            "progname": str(pname).strip() if pname else "",
            "genre": str(cell("genre")).strip() if cell("genre") else "",
            # durations kept in SECONDS internally; unit resolved per-show.
            "prog_dur_s": parse_duration_to_seconds(cell("pdur")),
            "brk_no": cell("brk"),
            "ad_dur_s": parse_duration_to_seconds(cell("dur")),
            # raw cell values, kept so the unit can be re-checked if needed.
            "prog_dur_raw": cell("pdur"),
            "ad_dur_raw": cell("dur"),
            "ad_st": parse_time_to_seconds(cell("adst")),
            "ad_et": parse_time_to_seconds(cell("adet")),
            "brand": str(cell("brand")).strip() if cell("brand") else "",
        }
        spots.append(spot)
    return spots


# ---------------------------------------------------------------------------
# The transform: spots -> show rows
# ---------------------------------------------------------------------------

def _brk_key(value):
    """Normalise a break number to an int when possible, else a string."""
    if value is None or value == "":
        return None
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return str(value).strip()


def _detect_duration_unit(spots):
    """
    Decide, ONCE for the whole file, what unit the per-ad DURATION column
    uses. Returns the number of seconds that one DURATION unit represents:
        1     -> DURATION is already in seconds
        60    -> DURATION is in minutes
        3600  -> DURATION is in hours
        86400 -> DURATION is an Excel fraction-of-a-day

    Method: a single television ad spot is, in the real world, almost always
    between ~3 and ~180 seconds long. We take the median raw DURATION value
    across every ad in the file and pick the unit under which that median
    falls into the plausible per-ad range. Deciding once, file-wide, avoids
    the per-show guessing that produced inconsistent results.

    A second, independent check uses each ad's own ADST->ADET timestamps
    (the ad's true length, unit-free); if DURATION strongly disagrees with
    those, the timestamp-derived unit wins.
    """
    raw_vals = []
    span_vals = []          # ad length in seconds from ADST->ADET
    for s in spots:
        rv = s.get("ad_dur_raw")
        if isinstance(rv, (int, float)) and rv > 0:
            raw_vals.append(float(rv))
        elif rv not in (None, ""):
            # string like "00:00:30" -> already unambiguous seconds
            secs = parse_time_to_seconds(str(rv))
            if secs:
                span_vals.append(float(secs))
        ast_, aet = s.get("ad_st"), s.get("ad_et")
        if ast_ is not None and aet is not None:
            d = aet - ast_
            if d < 0:
                d += 86400
            if 0 < d < 1800:        # sane single-ad span
                span_vals.append(float(d))

    def _median(xs):
        xs = sorted(xs)
        n = len(xs)
        if n == 0:
            return None
        return xs[n // 2] if n % 2 else (xs[n // 2 - 1] + xs[n // 2]) / 2.0

    # If DURATION column is entirely time strings, it is already seconds.
    if not raw_vals:
        return 1

    med = _median(raw_vals)
    candidates = [
        (1, "seconds"),
        (60, "minutes"),
        (3600, "hours"),
        (86400, "fraction-of-day"),
    ]
    # Plausible per-ad length: 3..180 seconds (allow a little slack).
    LOW, HIGH = 2.0, 600.0
    scored = []
    for unit_secs, _name in candidates:
        per_ad_seconds = med * unit_secs
        in_range = LOW <= per_ad_seconds <= HIGH
        # distance from the centre of the plausible band (log scale)
        import math as _m
        centre = _m.sqrt(LOW * HIGH)
        dist = abs(_m.log(per_ad_seconds / centre)) if per_ad_seconds > 0 else 1e9
        scored.append((not in_range, dist, unit_secs))
    scored.sort()
    chosen = scored[0][2]

    # Cross-check against ADST->ADET spans if we have enough of them.
    span_med = _median(span_vals)
    if span_med and len(span_vals) >= max(3, len(raw_vals) // 4):
        # what unit makes DURATION's median match the timestamp median?
        if med > 0:
            implied = span_med / med          # seconds per DURATION unit
            for unit_secs, _ in candidates:
                if 0.5 * unit_secs <= implied <= 2.0 * unit_secs:
                    chosen = unit_secs
                    break
    return chosen


def transform(spots):
    """
    Transform a flat list of ad-spot dicts into Zee-format show rows.
    Returns a list of dicts keyed by OUTPUT_COLS.
    """
    # --- Stage 1: decide the DURATION column's unit ONCE for the whole file -
    # The DURATION column uses one consistent unit throughout a file. We
    # detect it from all ads together, not per-show, so every channel and
    # every programme is treated identically.
    duration_unit_s = _detect_duration_unit(spots)

    # Convert each ad's DURATION to seconds using that single decision.
    # If DURATION was a time string, ad_dur_s already holds true seconds and
    # we keep it; numeric DURATION values are scaled by the detected unit.
    for s in spots:
        rv = s.get("ad_dur_raw")
        if isinstance(rv, (int, float)):
            s["ad_seconds"] = float(rv) * duration_unit_s
        else:
            # string / time value -> ad_dur_s is already seconds
            s["ad_seconds"] = s.get("ad_dur_s") or 0.0

    # --- Stage 2: group spots by channel + date + programme + prog start ----
    groups = {}
    order = []
    for s in spots:
        key = (s["channel"], s["date"], s["progname"], s["prog_st"])
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(s)

    show_rows = []
    for key in order:
        channel, d, progname, prog_st = key
        members = groups[key]
        first = members[0]
        prog_et = first["prog_et"]
        genre = first["genre"]

        # --- Programme duration -------------------------------------------
        # PROG ST and PROG ET are unambiguous timestamps, so the real
        # airtime is computed directly from them. This is authoritative and
        # does not depend on the file's PROG DUR column at all.
        if prog_st is not None and prog_et is not None:
            span = prog_et - prog_st
            if span < 0:                       # programme crosses midnight
                span += 86400
            prog_dur = span / 60.0             # minutes
        elif prog_et is None and prog_st is not None:
            prog_dur = 0.0
        else:
            prog_dur = 0.0

        # --- Stage 4: bucket spots by break number --------------------------
        breaks = {}
        for s in members:
            bk = _brk_key(s["brk_no"])
            if bk is None:
                continue
            breaks.setdefault(bk, []).append(s)

        # Sort break keys: numerics first (ascending), then any string keys.
        numeric_keys = sorted([k for k in breaks if isinstance(k, int)])
        string_keys = sorted([k for k in breaks if not isinstance(k, int)],
                              key=str)
        ordered_break_keys = numeric_keys + string_keys

        break_windows = []  # list of (start_sec, end_sec) per break, in order
        for bk in ordered_break_keys:
            bspots = breaks[bk]
            starts = [s["ad_st"] for s in bspots if s["ad_st"] is not None]
            ends = [s["ad_et"] for s in bspots if s["ad_et"] is not None]
            if not starts or not ends:
                continue
            break_windows.append((min(starts), max(ends)))

        # Sort break windows chronologically by start time.
        break_windows.sort(key=lambda w: w[0])

        n_breaks = len(break_windows)
        # Per-ad durations were all converted to seconds with one file-wide
        # unit decision (Stage 1). Sum them and convert to minutes.
        ad_dur = sum(s.get("ad_seconds", 0.0) for s in members) / 60.0
        content_dur = (prog_dur or 0) - ad_dur
        # Content duration can't sensibly be negative; clamp at 0 if the
        # source data is inconsistent.
        if content_dur < 0:
            content_dur = 0

        # --- Stage 5: segments ---------------------------------------------
        # A show with N breaks has N+1 segments (content stretches).
        segments = []  # minutes
        if prog_st is not None:
            if n_breaks == 0:
                if prog_et is not None:
                    segments.append((prog_et - prog_st) / 60.0)
            else:
                # Seg 1: prog start -> break 1 start
                segments.append((break_windows[0][0] - prog_st) / 60.0)
                # Seg n: end of break (n-1) -> start of break n
                for i in range(1, n_breaks):
                    segments.append(
                        (break_windows[i][0] - break_windows[i - 1][1]) / 60.0
                    )
                # Last Seg: end of last break -> prog end
                if prog_et is not None:
                    segments.append(
                        (prog_et - break_windows[-1][1]) / 60.0
                    )

        # --- Build the output row ------------------------------------------
        row = {col: "" for col in OUTPUT_COLS}
        row["WK CODE"] = wk_code(d)
        row["Week"] = wk_code(d)
        row["Day"] = DAY_NAMES[d.weekday()] if d else ""
        row["Date"] = d.strftime("%d/%m/%Y") if d else ""
        row["Start Time"] = seconds_to_hms(start_time_rounded(prog_st))
        row["PROGNAME"] = progname
        row["PROG ST"] = seconds_to_hms(prog_st)
        row["PROG ET"] = seconds_to_hms(prog_et)
        # Rounding rule:
        #   Prog Dur, Ad Dur, No.Of Breaks -> numeric, rounded to 1 decimal.
        #   everything else (Content Dur, segments, Mins Missing) -> whole
        #     numbers, no decimals.
        # Kept as real numbers (not strings) so sorting and Excel formulas
        # work normally.
        row["Content Dur"] = round(content_dur)
        row["Prog Dur"] = round(prog_dur, 1) if prog_dur else 0
        row["Ad Dur"] = round(ad_dur, 1)
        row["No.Of Breaks"] = round(n_breaks, 1)
        # Mins Missing filled in stage 6 below.

        # Break Start/End pairs (up to 18)
        for i, (bs, be) in enumerate(break_windows[:18], start=1):
            row[f"Ad ST {i}"] = seconds_to_hms(bs)
            row[f"Ad ET {i}"] = seconds_to_hms(be)

        # Segments: Seg 1..18 then Last Seg — whole numbers (no decimals)
        if segments:
            if n_breaks == 0:
                row["Seg 1"] = round(segments[0])
            else:
                # all but the last go into Seg 1..18
                for i, seg in enumerate(segments[:-1][:18], start=1):
                    row[f"Seg {i}"] = round(seg)
                row["Last Seg"] = round(segments[-1])

        # keep raw sort/filter keys for later (channel filtering, date
        # grouping, mins missing) — these never go into the output table.
        row["_channel"] = channel
        row["_date"] = d
        row["_prog_st"] = prog_st
        row["_prog_et"] = prog_et
        show_rows.append(row)

    # --- Stage 6: Mins Missing ---------------------------------------------
    # Per channel, in chronological order: this show's start - prev show's end.
    by_channel = {}
    for r in show_rows:
        by_channel.setdefault(r["_channel"], []).append(r)

    for channel, rows_for_ch in by_channel.items():
        rows_for_ch.sort(
            key=lambda r: (
                r["_date"] or date.min,
                r["_prog_st"] if r["_prog_st"] is not None else 0,
            )
        )
        prev_end = None
        prev_date = None
        for r in rows_for_ch:
            if prev_end is None or r["_prog_st"] is None:
                r["Mins Missing"] = ""
            elif prev_date != r["_date"]:
                # first show of a new date for this channel -> blank
                r["Mins Missing"] = ""
            else:
                gap = (r["_prog_st"] - prev_end) / 60.0
                r["Mins Missing"] = round(gap)
            prev_end = r["_prog_et"]
            prev_date = r["_date"]

    return show_rows


def transform_workbook_rows(rows):
    """Convenience: brand-ads sheet rows -> Zee-format show rows."""
    spots = read_brand_ads(rows)
    return transform(spots)