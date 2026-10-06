"""
topsheet_parser.py - lettura flessibile dei topsheet .xlsx

Invece di leggere le colonne "a posizione fissa", cerca le intestazioni
(SHOP NAME, TRADE NAME, CITY, PCS, VALUE, SEL, OC, PROD, PAYMENT TERMS,
RC DEL, MJE DEL, COMMENTS, ORDER N., ...) ovunque si trovino, quindi
funziona anche se i file hanno colonne diverse, righe di titolo in più,
righe di sezione ("EUR ACCOUNTS", "TOTAL USD") e totali.

parse_xlsx(file) -> (brand, stagione, DataFrame)
"""
import re

import openpyxl
import pandas as pd

OUT_COLS = ["cliente", "trade_name", "city", "qty", "pct", "valuta", "sel", "oc", "prod",
            "payment_terms", "producer_code", "rc_del", "mje_del", "comments",
            "pay_status", "ship_status"]
SEASON_RE = re.compile(r"\b([A-Z]{2}\d{2})\b")
SKIP_PREFIX = ("TOTAL", "GRAND", "SUBTOTAL", "SUB TOTAL", "TOT ")


def _txt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _label(v):
    return re.sub(r"\s+", " ", _txt(v).upper().replace("\t", " ")).strip()


def _num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    s = _txt(v).replace("€", "").replace("$", "").replace(" ", "")
    if not s or s.startswith("="):
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _currency(fmt):
    """'€' o '$' dal formato numerico della cella dell'importo (None se non riconoscibile)."""
    f = str(fmt or "").upper()
    if "€" in f or "EUR" in f:
        return "€"
    if "$" in f or "USD" in f:
        return "$"
    return None


def _flag(v):
    return _label(v) in {"Y", "YES", "SI", "SÌ", "X", "TRUE", "1", "OK", "✓", "V"}


def _title(ws):
    """Brand e stagione dal titolo (prima cella di testo in alto a sinistra)."""
    for row in ws.iter_rows(min_row=1, max_row=3, max_col=4):
        for c in row:
            t = _txt(c.value)
            if t:
                m = None
                for m in SEASON_RE.finditer(t.upper()):
                    pass
                if m:
                    return t[:m.start()].strip(" -_"), m.group(1)
                return t, ""
    return "", ""


def parse_xlsx(file):
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb.worksheets[0]
    brand, season = _title(ws)

    # 1) riga delle intestazioni principali: quella che contiene "TRADE NAME"
    head_row = trade_col = None
    for r in range(1, min(ws.max_row, 30) + 1):
        for c in range(1, min(ws.max_column, 40) + 1):
            if _label(ws.cell(r, c).value) == "TRADE NAME":
                head_row, trade_col = r, c
                break
        if head_row:
            break
    if not head_row:
        raise ValueError("Non trovo l'intestazione «TRADE NAME»: controlla che sia nel file.")

    # 2) prima riga di dati: dopo l'intestazione, con un trade name e non un totale
    def is_total(t):
        return _label(t).startswith(SKIP_PREFIX)

    first = None
    for r in range(head_row + 1, ws.max_row + 1):
        t = _txt(ws.cell(r, trade_col).value)
        if t and not is_total(t):
            first = r
            break
    if not first:
        raise ValueError("Non trovo nessuna riga di ordini sotto l'intestazione.")

    # 3) etichetta di ogni colonna = ultima cella di testo sopra i dati
    labels = {}
    for c in range(1, ws.max_column + 1):
        for r in range(1, first):
            v = _label(ws.cell(r, c).value)
            if v:
                labels[c] = v

    def find(pred, exclude=()):
        for c in sorted(labels):
            if c not in exclude and pred(labels[c]):
                return c
        return None

    # shop / città si cercano nella riga delle intestazioni principali (sotto potrebbero esserci righe di sezione)
    head = {c: _label(ws.cell(head_row, c).value) for c in range(1, ws.max_column + 1)}
    shop_col = next((c for c, l in head.items() if l.startswith("SHOP") or l in ("CLIENT", "CLIENTE", "CUSTOMER")), None)
    city_col = next((c for c, l in head.items() if l in ("CITY", "CITTA", "CITTÀ")), None)
    qty_col = find(lambda l: l in ("PCS", "QTY", "QTÀ", "QTA", "PIECES", "QUANTITY"))
    val_col = find(lambda l: l in ("VALUE", "AMOUNT", "IMPORTO", "TOTAL VALUE"))
    if qty_col is None and city_col:
        qty_col = city_col + 1          # nei file senza etichette PCS/VALUE: subito dopo la città
    if val_col is None and qty_col:
        val_col = qty_col + 1
    used = {c for c in (trade_col, shop_col, city_col, qty_col, val_col) if c}

    sel_col = find(lambda l: l == "SEL", used)
    oc_col = find(lambda l: l == "OC", used)
    prod_col = find(lambda l: l == "PROD", used)
    pay_col = find(lambda l: l.startswith("PAYMENT"), used)
    rc_col = find(lambda l: l == "RC DEL" or l.startswith("DEL TO"), used)
    mje_col = find(lambda l: l == "MJE DEL", used)
    com_col = find(lambda l: l.startswith("COMMENT"), used)
    code_col = find(lambda l: l.startswith("ORDER") or (season and l == season.upper()),
                    used | {sel_col, oc_col, prod_col, pay_col, rc_col, mje_col, com_col})
    # se più colonne si chiamano COMMENTS (banner + colonna vera) vale l'ultima, già garantito da labels

    def cell(r, c):
        return ws.cell(r, c).value if c else None

    rows = []
    ctx = "€"
    for r in range(head_row + 1, ws.max_row + 1):
        trade = _txt(cell(r, trade_col))
        if not trade:
            # riga di sezione (es. "USD ACCOUNTS"): vale come valuta di default per le righe seguenti
            head_txt = " ".join(_label(ws.cell(r, c).value) for c in (1, 2))
            if not head_txt.startswith(SKIP_PREFIX):
                if "USD" in head_txt or "$" in head_txt:
                    ctx = "$"
                elif "EUR" in head_txt or "€" in head_txt:
                    ctx = "€"
        if r < first or not trade or is_total(trade):
            continue
        qty, val = _num(cell(r, qty_col)), _num(cell(r, val_col))
        if qty is None and val is None:
            continue
        shop = _txt(cell(r, shop_col)) or trade
        terms = _txt(cell(r, pay_col))
        rows.append({
            "cliente": shop,
            "trade_name": trade,
            "city": _txt(cell(r, city_col)),
            "qty": qty,
            "pct": val,
            "valuta": _currency(ws.cell(r, val_col).number_format) or ctx,
            "sel": _flag(cell(r, sel_col)),
            "oc": _flag(cell(r, oc_col)),
            "prod": _flag(cell(r, prod_col)),
            "payment_terms": terms,
            "producer_code": _txt(cell(r, code_col)),
            "rc_del": _txt(cell(r, rc_col)),
            "mje_del": _txt(cell(r, mje_col)),
            "comments": _txt(cell(r, com_col)),
            # stesso comportamento del vecchio lettore: "DP PAID" nei termini = acconto già pagato
            "pay_status": "acconto_pagato" if "DP PAID" in terms.upper() else "da_pagare",
            "ship_status": "da_spedire",
        })
    if not rows:
        raise ValueError("Nessun ordine valido trovato nel file.")
    return brand, season, pd.DataFrame(rows, columns=OUT_COLS)
