"""
app.py — Flask backend for SHOWLOG / DataHub.

Run:  python app.py
Then open http://localhost:5000 in a browser and sign in.

On a shared server, run behind a production WSGI server (waitress):
    python -m waitress --host=0.0.0.0 --port=5000 app:app
"""

import io
import os
import datetime
import traceback
import secrets
from functools import wraps

from flask import (Flask, render_template, request, jsonify,
                   send_file, abort, session, redirect, url_for)

import database as db
import auth
from excelio import (read_workbook_rows, build_export,
                     build_multi_sheet_export)
from transform import (transform_workbook_rows, OUTPUT_COLS, display_label,
                       SHOW_INFO_COLS, BREAK_COLS, SEG_COLS)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB upload cap

# Session secret: persisted to a file so logins survive a server restart.
_SECRET_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            ".secret_key")
if os.path.exists(_SECRET_FILE):
    with open(_SECRET_FILE) as f:
        app.secret_key = f.read().strip()
else:
    app.secret_key = secrets.token_hex(32)
    with open(_SECRET_FILE, "w") as f:
        f.write(app.secret_key)

app.permanent_session_lifetime = datetime.timedelta(hours=12)

db.init_db()

ALLOWED_EXT = {".xlsx", ".xlsm", ".xls"}


# ---------------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------------

def login_required(fn):
    """Decorator: redirect to /login (pages) or 401 (API) if not signed in."""
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("user"):
            if request.path.startswith("/api/"):
                return jsonify(error="Not signed in."), 401
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper


@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("user"):
        return redirect(url_for("index"))
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        if auth.verify(username, password):
            session.permanent = True
            session["user"] = username.strip()
            return redirect(url_for("index"))
        return render_template("login.html",
                               error="Invalid username or password.")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# Page
# ---------------------------------------------------------------------------

@app.route("/")
@login_required
def index():
    return render_template(
        "index.html",
        show_info_cols=SHOW_INFO_COLS,
        break_cols=BREAK_COLS,
        seg_cols=SEG_COLS,
        all_cols=OUTPUT_COLS,
        col_labels={c: display_label(c) for c in OUTPUT_COLS},
        channel_buttons=db.CHANNEL_BUTTONS,
        current_user=session.get("user"),
    )


# ---------------------------------------------------------------------------
# Upload + accumulate
# ---------------------------------------------------------------------------

@app.route("/api/upload", methods=["POST"])
@login_required
def api_upload():
    files = request.files.getlist("files")
    if not files:
        return jsonify(error="No files received."), 400

    results = []
    for f in files:
        name = f.filename or "unnamed.xlsx"
        ext = "." + name.rsplit(".", 1)[-1].lower() if "." in name else ""
        if ext not in ALLOWED_EXT:
            results.append({"filename": name, "ok": False,
                            "error": f"Unsupported file type ({ext})."})
            continue
        try:
            data = f.read()
            rows = read_workbook_rows(io.BytesIO(data))
            show_rows = transform_workbook_rows(rows)
            if not show_rows:
                results.append({"filename": name, "ok": False,
                                "error": "No shows found in file."})
                continue
            uploaded_at = datetime.datetime.now().isoformat(timespec="seconds")
            db.add_file(name, uploaded_at, show_rows)
            results.append({"filename": name, "ok": True,
                            "shows": len(show_rows)})
        except Exception as e:  # noqa: BLE001
            traceback.print_exc()
            results.append({"filename": name, "ok": False, "error": str(e)})

    return jsonify(results=results)


# ---------------------------------------------------------------------------
# Data
# ---------------------------------------------------------------------------

@app.route("/api/data")
@login_required
def api_data():
    """Return all accumulated shows + file list + channel map + channels."""
    shows = db.get_all_shows()
    return jsonify(
        shows=shows,
        files=db.list_files(),
        channel_map=db.get_channel_map(),
        channels=db.distinct_channels(),
        columns={
            "show_info": SHOW_INFO_COLS,
            "ad_breaks": BREAK_COLS,
            "segments": SEG_COLS,
            "all": OUTPUT_COLS,
        },
    )


@app.route("/api/files/<int:file_id>", methods=["DELETE"])
@login_required
def api_delete_file(file_id):
    db.delete_file(file_id)
    return jsonify(ok=True)


@app.route("/api/clear", methods=["POST"])
@login_required
def api_clear():
    db.clear_all()
    return jsonify(ok=True)


# ---------------------------------------------------------------------------
# Channel mapper
# ---------------------------------------------------------------------------

@app.route("/api/channel-map", methods=["GET", "POST"])
@login_required
def api_channel_map():
    if request.method == "POST":
        mapping = request.get_json(silent=True) or {}
        db.set_channel_map(mapping)
        return jsonify(ok=True, channel_map=db.get_channel_map())
    return jsonify(channel_map=db.get_channel_map(),
                   channels=db.distinct_channels())


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

@app.route("/api/export")
@login_required
def api_export():
    """
    Export shows.
      * format=xlsx -> ONE workbook with an "All" sheet plus one sheet per
                       mapped channel that has data.
      * format=csv  -> a single flat table (CSV cannot hold multiple sheets).
    """
    fmt = request.args.get("format", "xlsx").lower()
    if fmt not in ("xlsx", "csv"):
        abort(400)

    shows = db.get_all_shows()
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")

    if fmt == "csv":
        # CSV: optional single-channel filter, single table.
        channel_value = request.args.get("channel")
        if channel_value:
            shows = [s for s in shows if s.get("_channel") == channel_value]
        data, mimetype, ext = build_export(shows, fmt="csv")
    else:
        # XLSX: multi-sheet workbook. "All" + one sheet per channel button.
        # All 7 buttons always get a sheet, even if empty.
        channel_map = db.get_channel_map()      # {button: raw CHNLNAME}
        channel_sheets = []
        for button in db.CHANNEL_BUTTONS:
            raw = channel_map.get(button) or ""
            rows = [s for s in shows if raw and s.get("_channel") == raw]
            channel_sheets.append((button, rows))
        data, mimetype, ext = build_multi_sheet_export(shows, channel_sheets)

    fname = f"SHOWLOG_{stamp}.{ext}"
    return send_file(io.BytesIO(data), mimetype=mimetype,
                     as_attachment=True, download_name=fname)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)