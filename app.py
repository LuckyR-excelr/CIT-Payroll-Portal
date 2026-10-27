"""CIT Payroll Portal - run:  python app.py   then open http://127.0.0.1:5000"""
import datetime as dt
import io
import os
import secrets
import uuid
from decimal import Decimal

from flask import (Flask, Response, abort, flash, jsonify, redirect, render_template, request, send_file,
                   url_for)
from werkzeug.utils import secure_filename

from payroll import store
from payroll.columns import EDITABLE, NUMERIC_KEYS
from payroll.config import ALLOWED_EXT, DEFAULT_MAPPINGS, MAX_UPLOAD_MB
from payroll.engine import (checks, compute, generate, is_expense, readiness, summary, to_baseline,
                            _flag_resolved)
from payroll.exporter import build_workbook
from payroll.readers import MappingError, merge_mappings, read_baseline, read_table
from payroll.util import D, ZERO, num, pay_period, q2, to_date, trunc3

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
app.secret_key = os.environ.get("PORTAL_SECRET") or secrets.token_hex(16)


# ------------------------------------------------------------------ optional password (set PORTAL_PASSWORD)
@app.before_request
def _auth():
    pw = os.environ.get("PORTAL_PASSWORD")
    if not pw:
        return None
    a = request.authorization
    if a and secrets.compare_digest(a.password or "", pw):
        return None
    return Response("Login required", 401, {"WWW-Authenticate": 'Basic realm="CIT Payroll Portal"'})


# ------------------------------------------------------------------ helpers
def _ext_ok(f):
    return f and f.filename and os.path.splitext(f.filename)[1].lower() in ALLOWED_EXT


def _next_run_date(today=None):
    today = today or dt.date.today()
    d = dt.date(today.year, today.month, 15)
    if today > d:
        d = dt.date(today.year + (today.month == 12), (today.month % 12) + 1, 15)
    return d


def _s(v, kind="money"):
    if v is None:
        return ""
    if isinstance(v, Decimal):
        return str(q2(v)) if kind == "money" else format(v.normalize(), "f")
    return str(v)


def _flag_view(f, state):
    d = dict(f)
    d["resolved"] = bool(f["severity"] == "error" and _flag_resolved(f, state))
    return d


def _row_view(r, c, flags_by_uid):
    return {
        "uid": r["uid"], "block": r["block"], "ptype": r.get("ptype") or ("MPS" if r["block"] == "mps" else "Regular"),
        "emp_id": r.get("emp_id") or "", "name": r.get("name") or "", "vendor": r.get("vendor") or "",
        "term": r.get("term") or "", "mode": r.get("mode") or "", "pri_sec": r.get("pri_sec") or "",
        "hours": _s(r.get("hours"), "n"), "inv_hours": r.get("inv_hours_text") or _s(r.get("inv_hours"), "n"),
        "inv_rate": _s(r.get("inv_rate"), "n"), "cons_rate": _s(r.get("cons_rate"), "n"), "cons_pct": _s(r.get("cons_pct"), "n"),
        "bonus": _s(r.get("bonus"), "n"), "deduct": _s(r.get("deduct"), "n"), "hi": _s(r.get("hi"), "n"),
        "expenses": _s(r.get("expenses"), "n"), "ae_override": _s(r.get("ae_override"), "n"),
        "ao_override": _s(r.get("ao_override"), "n"), "comments": r.get("comments") or "",
        "excluded": bool(r.get("excluded")), "is_expense": is_expense(r),
        "S": _s(c["S"], "n"), "W": _s(c["W"]), "AB": _s(c["AB"]), "AC": _s(c["AC"]), "AE": _s(c["AE"], "n"),
        "AF": _s(trunc3(c["AF"]) if c["AF"] is not None else None, "n"), "AO": _s(c["AO"], "n"),
        "flag_ids": flags_by_uid.get(r["uid"], []),
    }


def build_payload(mid, state):
    rd = readiness(state)
    comp = {r["uid"]: compute(r) for r in state["rows"]}
    fb = {}
    for f in state["flags"]:
        if f.get("uid"):
            fb.setdefault(f["uid"], []).append(f["id"])
    s = summary(state)
    return {
        "month": {"id": mid, "run_date": state["run_date"], "pay_period": state["pay_period"],
                  "status": state.get("status", "draft"), "files": state.get("files", {})},
        "rows": [_row_view(r, comp[r["uid"]], fb) for r in state["rows"]],
        "flags": [_flag_view(f, state) for f in state["flags"]],
        "errors": rd["errors"], "checks": rd["checks"],
        "unacked": len(rd["unacked"]), "can_download": rd["can_download"], "can_close": rd["can_close"],
        "summary": {k: (_s(v) if isinstance(v, Decimal) else v) for k, v in s.items()},
        "carrier_total": _s(state.get("carrier_total")),
    }


def _load(mid):
    st = store.get_month(mid)
    if st is None:
        abort(404)
    return st


def _editable(state):
    if state.get("status") == "closed":
        return jsonify(error="This month is closed and read-only."), 409
    return None


# ------------------------------------------------------------------ pages
@app.route("/")
def dashboard():
    base = store.get_baseline()
    months = []
    for m in store.month_ids():
        st = store.get_month(m)
        s = summary(st)
        months.append({"id": m, "run_date": st["run_date"], "pay_period": st["pay_period"], "status": st.get("status", "draft"),
                       "gross": _s(s["regular_gross"] + s["mps_gross"]), "rows": s["rows"]})
    return render_template("dashboard.html", baseline=base, months=months, next_run=_next_run_date().isoformat(), page="home")


@app.post("/baseline")
def upload_baseline():
    f = request.files.get("file")
    if not _ext_ok(f) or f.filename.lower().endswith(".csv"):
        flash("Upload the previous CIT payroll as an .xlsx / .xlsm file.", "error")
        return redirect(url_for("dashboard"))
    tmp = store.path("uploads", f"{uuid.uuid4().hex}_{secure_filename(f.filename)}")
    f.save(tmp)
    try:
        b = read_baseline(tmp)
    except MappingError as e:
        flash(f"Could not read that payroll file: {e}", "error")
        return redirect(url_for("dashboard"))
    b["summary"]["source_file"] = f.filename
    b["summary"]["imported"] = dt.datetime.now().isoformat(timespec="minutes")
    store.save_baseline(b, dt.datetime.now().strftime("%Y%m%d%H%M%S"))
    n_one = sum(1 for r in b["rows"] if r["one_time"])
    flash(f"Baseline loaded: {len(b['rows'])} rows from '{f.filename}' (pay period {b['summary'].get('pay_period')}). "
          f"{n_one} one-time row(s) will not be carried forward.", "ok")
    return redirect(url_for("dashboard"))


@app.route("/month/new", methods=["GET", "POST"])
def new_month():
    base = store.get_baseline()
    if not base:
        flash("Load last month's payroll first (Dashboard, step 1).", "error")
        return redirect(url_for("dashboard"))
    if request.method == "GET":
        return render_template("new_month.html", next_run=_next_run_date().isoformat(), page="new")
    try:
        run_date = to_date(request.form["run_date"])
    except Exception:
        flash("Choose a valid run date.", "error")
        return redirect(url_for("new_month"))
    mid = f"{run_date:%Y-%m}"
    existing = store.get_month(mid)
    if existing and existing.get("status") == "closed":
        flash(f"{mid} is already closed. It cannot be regenerated.", "error")
        return redirect(url_for("dashboard"))
    fi, fh = request.files.get("invoices"), request.files.get("health")
    if not (_ext_ok(fi) and _ext_ok(fh)):
        flash("Upload both the CIT invoice sheet and the health-insurance file (.xlsx or .csv).", "error")
        return redirect(url_for("new_month"))
    pi = store.input_path(mid, "invoices_" + secure_filename(fi.filename))
    ph = store.input_path(mid, "health_" + secure_filename(fh.filename))
    fi.save(pi)
    fh.save(ph)
    maps = store.get_mappings()
    try:
        invoices, im = read_table(pi, "invoices", maps, request.form.get("invoice_sheet") or None)
        hi, hm = read_table(ph, "health", maps, request.form.get("health_sheet") or None)
    except MappingError as e:
        found = ", ".join(e.headers[:20]) or "(none)"
        flash(f"{e} Columns found in the first sheet: {found}. Sheets: {', '.join(e.sheets)}. "
              "Add the real column names in Settings -> Column names, then try again.", "error")
        return redirect(url_for("new_month"))
    state = generate(base, invoices, hi, store.get_rules(), run_date)
    state.update({"status": "draft", "created": dt.datetime.now().isoformat(timespec="minutes"),
                  "files": {"invoices": fi.filename, "health": fh.filename, "invoice_columns": im["columns"],
                            "invoice_sheet": im["sheet"], "health_sheet": hm["sheet"], "health_columns": hm["columns"],
                            "invoice_lines": im["records"], "health_lines": hm["records"]}})
    store.save_month(mid, state)
    flash(f"Generated {len(state['rows'])} payroll rows from {im['records']} invoice lines and {hm['records']} insurance lines.", "ok")
    return redirect(url_for("review", mid=mid))


@app.route("/month/<mid>")
def review(mid):
    st = _load(mid)
    return render_template("review.html", mid=mid, state_status=st.get("status", "draft"), page="review")


@app.route("/settings", methods=["GET", "POST"])
def settings():
    if request.method == "POST":
        maps = {"invoices": {}, "health": {}}
        for kind in maps:
            for field in DEFAULT_MAPPINGS[kind]:
                raw = request.form.get(f"map_{kind}_{field}", "")
                names = [x.strip() for x in raw.split(",") if x.strip()]
                defaults = [d for d in DEFAULT_MAPPINGS[kind][field]]
                extra = [n for n in names if n.lower() not in [d.lower() for d in defaults]]
                if extra:
                    maps[kind][field] = extra
        store.save_mappings(maps)
        rules = {"hi_employer_paid": {}, "split": {}}
        for i, a in zip(request.form.getlist("hi_id"), request.form.getlist("hi_amt")):
            if i.strip() and D(a) is not None:
                rules["hi_employer_paid"][i.strip()] = str(D(a))
        for i, h in zip(request.form.getlist("sp_id"), request.form.getlist("sp_hrs")):
            if i.strip() and D(h) is not None:
                rules["split"][i.strip()] = {"regular_hours": str(D(h))}
        store.save_rules(rules)
        flash("Settings saved. They apply to the next month you generate.", "ok")
        return redirect(url_for("settings"))
    custom = store.get_mappings()
    merged = merge_mappings(custom)
    view = {k: {f: {"defaults": DEFAULT_MAPPINGS[k][f], "extra": (custom.get(k, {}) or {}).get(f, [])} for f in DEFAULT_MAPPINGS[k]}
            for k in DEFAULT_MAPPINGS}
    return render_template("settings.html", view=view, rules=store.get_rules(), page="settings")


# ------------------------------------------------------------------ JSON API used by the review grid
@app.get("/api/month/<mid>")
def api_month(mid):
    return jsonify(build_payload(mid, _load(mid)))


def _num_field(key, raw):
    if raw is None or str(raw).strip() == "":
        return None
    v = D(raw)
    if v is None:
        raise ValueError(f"'{raw}' is not a number")
    if key == "cons_pct" and v > Decimal("1.5"):
        v = v / 100
    return v


@app.post("/api/month/<mid>/row/<uid>")
def api_edit_row(mid, uid):
    st = _load(mid)
    err = _editable(st)
    if err:
        return err
    row = next((r for r in st["rows"] if r["uid"] == uid), None)
    if not row:
        abort(404)
    for k, v in (request.get_json(force=True) or {}).items():
        if k not in EDITABLE:
            return jsonify(error=f"Field '{k}' is not editable."), 400
        try:
            if k in NUMERIC_KEYS:
                row[k] = _num_field(k, v)
            elif k == "block":
                if v not in ("regular", "mps"):
                    raise ValueError("block must be regular or mps")
                row[k] = v
            else:
                row[k] = str(v).strip()
        except ValueError as e:
            return jsonify(error=str(e)), 400
    store.save_month(mid, st)
    return jsonify(build_payload(mid, st))


@app.post("/api/month/<mid>/row/<uid>/exclude")
def api_exclude(mid, uid):
    st = _load(mid)
    err = _editable(st)
    if err:
        return err
    row = next((r for r in st["rows"] if r["uid"] == uid), None)
    if not row:
        abort(404)
    row["excluded"] = not row.get("excluded")
    store.save_month(mid, st)
    return jsonify(build_payload(mid, st))


@app.post("/api/month/<mid>/row/<uid>/delete")
def api_delete(mid, uid):
    st = _load(mid)
    err = _editable(st)
    if err:
        return err
    st["rows"] = [r for r in st["rows"] if r["uid"] != uid]
    st["flags"] = [f for f in st["flags"] if f.get("uid") != uid]
    store.save_month(mid, st)
    return jsonify(build_payload(mid, st))


@app.post("/api/month/<mid>/row")
def api_add(mid):
    st = _load(mid)
    err = _editable(st)
    if err:
        return err
    body = request.get_json(force=True) or {}
    start, end = pay_period(to_date(st["run_date"]))
    row = {k: None for k in ("emp_id", "name", "vendor", "term", "mode", "pri_sec", "hours", "inv_hours", "inv_rate",
                             "cons_rate", "cons_pct", "bonus", "deduct", "hi", "expenses", "paid_exp", "lca")}
    row.update({"uid": uuid.uuid4().hex[:8], "block": body.get("block", "regular"), "comments": "", "excluded": False,
                "pay_period": st["pay_period"], "pp_start": start.isoformat(), "pp_end": end.isoformat(),
                "pp_month": (to_date(st["run_date"]) - dt.timedelta(days=1)).isoformat(), "ae_override": None,
                "ao_override": None, "parent_uid": None, "hi_carrier": None, "hi_employer": None, "inv_hours_text": None})
    for k in ("name", "emp_id", "vendor", "term", "mode"):
        row[k] = str(body.get(k, "")).strip()
    for k in ("hours", "cons_rate", "cons_pct", "inv_rate"):
        try:
            row[k] = _num_field(k, body.get(k))
        except ValueError as e:
            return jsonify(error=str(e)), 400
    row["inv_hours"] = row["hours"]
    st["rows"].append(row)
    store.save_month(mid, st)
    return jsonify(build_payload(mid, st))


@app.post("/api/month/<mid>/flag/<fid>")
def api_flag(mid, fid):
    st = _load(mid)
    err = _editable(st)
    if err:
        return err
    f = next((x for x in st["flags"] if x["id"] == fid), None)
    if not f:
        abort(404)
    body = request.get_json(force=True) or {}
    if body.get("apply") and f.get("suggested"):
        row = next((r for r in st["rows"] if r["uid"] == f.get("uid")), None)
        if row:
            for k, v in f["suggested"].items():
                row[k] = D(v) if k in NUMERIC_KEYS else v
        f["acked"] = True
    else:
        f["acked"] = bool(body.get("ack", True))
    store.save_month(mid, st)
    return jsonify(build_payload(mid, st))


@app.post("/api/month/<mid>/ack_all")
def api_ack_all(mid):
    st = _load(mid)
    err = _editable(st)
    if err:
        return err
    for f in st["flags"]:
        if f["severity"] == "warn":
            f["acked"] = True
    store.save_month(mid, st)
    return jsonify(build_payload(mid, st))


@app.get("/month/<mid>/download")
def download(mid):
    st = _load(mid)
    rd = readiness(st)
    draft = not rd["can_download"]
    if draft and request.args.get("draft") != "1":
        flash("Unresolved errors - resolve them, or download a clearly marked DRAFT.", "error")
        return redirect(url_for("review", mid=mid))
    wb = build_workbook(st, rd, {"Invoice sheet": st.get("files", {}).get("invoices"),
                                 "Health insurance file": st.get("files", {}).get("health"),
                                 "Generated by": "CIT Payroll Portal"}, draft=draft)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    run = to_date(st["run_date"])
    name = f"CIT_payroll_{run:%b-%d-%Y}{'_DRAFT' if draft else ''}.xlsx"
    return send_file(buf, as_attachment=True, download_name=name,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/month/<mid>/close")
def close_month(mid):
    st = _load(mid)
    rd = readiness(st)
    if st.get("status") == "closed":
        return redirect(url_for("review", mid=mid))
    if not rd["can_close"]:
        flash("Cannot close: there are errors, unacknowledged flags or failing checks.", "error")
        return redirect(url_for("review", mid=mid))
    store.save_baseline(to_baseline(st), f"close_{mid}")
    st["status"] = "closed"
    st["closed"] = dt.datetime.now().isoformat(timespec="minutes")
    store.save_month(mid, st)
    flash(f"{mid} closed. Next month will start from this payroll automatically - just upload the new invoice and health files.", "ok")
    return redirect(url_for("dashboard"))


@app.post("/month/<mid>/reopen_draft")
def delete_draft(mid):
    st = _load(mid)
    if st.get("status") == "closed":
        flash("Closed months cannot be deleted.", "error")
    else:
        store.delete_month(mid)
        flash(f"Draft {mid} deleted.", "ok")
    return redirect(url_for("dashboard"))


if __name__ == "__main__":
    app.run(host=os.environ.get("HOST", "127.0.0.1"), port=int(os.environ.get("PORT", "5000")), debug=os.environ.get("DEBUG") == "1")
