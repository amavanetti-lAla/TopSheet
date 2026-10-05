import json
import re
import time
import uuid
from datetime import datetime

import gspread
import pandas as pd
import streamlit as st

from core import COLS, PAY, SHIP, eur, make_pdf, num, parse_xlsx

st.set_page_config(page_title="Topsheet", layout="wide")

ANA = {"Brand": ["name"], "Stagioni": ["brand", "code"], "Clienti": ["name"]}
LOGCOLS = ["quando", "chi", "id", "trade_name", "campo", "prima", "dopo"]
EDIT = ["cliente", "trade_name", "city", "qty", "pct", "sel", "oc", "prod", "payment_terms", "producer_code",
        "rc_del", "mje_del", "comments", "pay_status", "pay_paid", "pay_due", "ship_status", "ship_date"]
# campi che un reimport può aggiornare (stato pagamento e spedizione non vengono mai toccati)
UPD = ["cliente", "qty", "pct", "sel", "oc", "prod", "payment_terms", "producer_code", "rc_del", "mje_del",
       "comments"]
VIEWS = ["Panoramica", "Ordini", "Importa", "Anagrafiche", "PDF", "Log"]


# ---------- Google Sheet ----------
@st.cache_resource
def book():
    if "gcp_json" in st.secrets:  # metodo semplice: contenuto del file .json incollato così com'è
        info = json.loads(st.secrets["gcp_json"])
    else:
        info = dict(st.secrets["gcp_service_account"])
    gc = gspread.service_account_from_dict(info)
    return gc.open_by_key(st.secrets["sheet_id"])


@st.cache_resource
def sheet(name, cols):
    """Apre la scheda una sola volta (e la riusa), per non superare i limiti di Google."""
    try:
        return book().worksheet(name)
    except gspread.WorksheetNotFound:
        w = book().add_worksheet(name, 1000, len(cols))
        w.update([list(cols)], value_input_option="RAW")
        return w


def with_retry(fn):
    for attempt in range(3):
        try:
            return fn()
        except gspread.exceptions.APIError as e:
            if attempt == 2 or "429" not in str(e):
                raise
            time.sleep(8 * (attempt + 1))  # limite di Google raggiunto: aspetta e riprova


def _read(name, cols):
    """Lettura diretta dal foglio (senza cache)."""
    w = sheet(name, tuple(cols))
    rec = with_retry(lambda: w.get_all_records(numericise_ignore=["all"], default_blank=""))
    df = pd.DataFrame(rec)
    for c in cols:
        if c not in df:
            df[c] = ""
    return df[list(cols)]


@st.cache_data(ttl=120, show_spinner=False)
def read(name, cols):
    return _read(name, cols)


def write(name, df, cols):
    w = sheet(name, tuple(cols))
    values = [list(cols)] + df[list(cols)].fillna("").astype(str).values.tolist()
    with_retry(lambda: w.resize(rows=len(values) + 100))  # la scheda cresce (o si riduce) secondo i dati
    with_retry(w.clear)
    with_retry(lambda: w.update(values, value_input_option="RAW"))
    st.cache_data.clear()  # le prossime letture prendono i dati aggiornati


def backup_orders():
    """Prima di ogni salvataggio copia lo stato attuale degli ordini in 'Backup_Ordini'
    (sempre l'ultimo stato) e in 'Backup_AAAAMMGG' (il primo salvataggio del giorno; si tengono gli ultimi 14)."""
    try:
        vals = with_retry(lambda: sheet("Ordini", tuple(COLS)).get_all_values())
        if len(vals) < 2:
            return

        def put(name, only_new=False):
            try:
                w = book().worksheet(name)
            except gspread.WorksheetNotFound:
                w = book().add_worksheet(name, len(vals) + 20, len(vals[0]))
            else:
                if only_new:
                    return False
                with_retry(lambda: w.resize(rows=len(vals) + 20, cols=len(vals[0])))
                with_retry(w.clear)
            with_retry(lambda: w.update(vals, value_input_option="RAW"))
            return True

        put("Backup_Ordini")
        if put(f"Backup_{datetime.now():%Y%m%d}", only_new=True):
            old = sorted(w.title for w in book().worksheets() if re.fullmatch(r"Backup_\d{8}", w.title))
            for t in old[:-14]:
                book().del_worksheet(book().worksheet(t))
    except Exception as e:
        st.warning(f"Backup non riuscito (il salvataggio prosegue): {e}")


def log(rows):
    if rows:
        w = sheet("Log", tuple(LOGCOLS))
        with_retry(lambda: w.append_rows(rows, value_input_option="RAW"))


def load_orders(fresh=False):
    df = (_read if fresh else read)("Ordini", tuple(COLS))
    for c in ["sel", "oc", "prod"]:
        df[c] = df[c].astype(str).str.upper().isin(["TRUE", "Y", "1"])
    for c in ["qty", "pct", "pay_paid"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in ["pay_due", "ship_date"]:
        df[c] = pd.to_datetime(df[c], errors="coerce").apply(lambda x: x.date() if pd.notna(x) else None)
    df["pay_status"] = df["pay_status"].where(df["pay_status"].isin(PAY), PAY[0])
    df["ship_status"] = df["ship_status"].where(df["ship_status"].isin(SHIP), SHIP[0])
    return df


def save_orders(df):
    backup_orders()
    out = df.copy()
    for c in ["pay_due", "ship_date"]:
        out[c] = out[c].apply(lambda x: x.isoformat() if hasattr(x, "isoformat") and pd.notna(x) else "")
    for c in ["sel", "oc", "prod"]:
        out[c] = out[c].map(lambda x: "TRUE" if x else "FALSE")
    for c in ["qty", "pct", "pay_paid"]:
        out[c] = out[c].map(num)
    write("Ordini", out, COLS)


def norm(x):
    return "" if x is None or (not isinstance(x, str) and pd.isna(x)) else str(x)


def same(a, b):
    """Confronto tra valori: 13 e 13.0 sono uguali."""
    a, b = norm(a), norm(b)
    try:
        return float(a) == float(b)
    except ValueError:
        return a == b


def season_key(code):
    """SS26 < FW26 < SS27 ...: ordina per anno, poi primavera/estate prima di autunno/inverno."""
    c = str(code).upper().strip()
    m = re.search(r"(\d{2})\s*$", c)
    year = int(m.group(1)) if m else 0
    half = 0 if c.startswith("PS") else 1 if c.startswith(("SS", "S", "PE", "RE")) else 2
    return (year, half)


def money(df):
    """Per ogni ordine: importo, incassato e ancora da incassare (un ordine 'saldato' vale come incassato per intero)."""
    imp = df["pct"].fillna(0)
    saldato = df["pay_status"] == "saldato"
    inc = df["pay_paid"].fillna(0).where(~saldato, imp)
    return imp, inc, (imp - inc).clip(lower=0)


# ---------- Dati ----------
orders = load_orders()
brands = read("Brand", tuple(ANA["Brand"]))["name"].tolist()
seasons = read("Stagioni", tuple(ANA["Stagioni"]))
clients = read("Clienti", tuple(ANA["Clienti"]))["name"].tolist()

# Apertura di un ordine dalla Panoramica: va fatto prima di creare i widget
if "_open" in st.session_state:
    b_, s_ = st.session_state.pop("_open")
    st.session_state["brand_sel"], st.session_state["season_sel"] = b_, s_
    st.session_state["view"] = "Ordini"

with st.sidebar:
    st.header("Topsheet")
    who = st.text_input("Il tuo nome", key="who", help="Viene registrato come autore di ogni modifica.")
    brand = st.selectbox("Brand", brands, key="brand_sel") if brands else None
    opts = (sorted(seasons[seasons["brand"] == brand]["code"].tolist(), key=season_key, reverse=True)
            if brand else [])
    season = st.selectbox("Stagione", opts, key="season_sel") if opts else None
    if st.button("Ricarica dati"):
        st.cache_data.clear()
        st.session_state["edv"] = st.session_state.get("edv", 0) + 1  # scarta le modifiche non salvate
        st.rerun()

view = st.radio("Vista", VIEWS, horizontal=True, key="view", label_visibility="collapsed")

# ---------- Panoramica ----------
if view == "Panoramica":
    if orders.empty:
        st.info("Nessun topsheet caricato: usa la scheda Importa.")
    else:
        o = orders.copy()
        o["importo"], o["incassato"], o["da_incassare"] = money(o)
        saldato = o["pay_status"] == "saldato"
        o["da_saldare"] = ~saldato
        o["scaduti"] = ~saldato & o["pay_due"].notna() & (o["pay_due"] < datetime.now().date())
        o["da_spedire"] = o["ship_status"] == "da_spedire"
        g = (o.groupby(["brand", "stagione"])
             .agg(ordini=("id", "count"), qty=("qty", "sum"), importo=("importo", "sum"),
                  incassato=("incassato", "sum"), da_incassare=("da_incassare", "sum"),
                  da_saldare=("da_saldare", "sum"), scaduti=("scaduti", "sum"),
                  da_spedire=("da_spedire", "sum"))
             .reset_index())
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Importo totale", eur(g["importo"].sum()))
        m2.metric("Incassato", eur(g["incassato"].sum()))
        m3.metric("Da incassare", eur(g["da_incassare"].sum()))
        m4.metric("Pagamenti scaduti", int(g["scaduti"].sum()))
        recenti = st.toggle("Dalla stagione più recente", value=True)
        g["_k"] = g["stagione"].map(season_key)
        g = (g.sort_values(["_k", "brand"], ascending=[not recenti, True])
              .drop(columns="_k").reset_index(drop=True))
        st.caption("Clicca su una riga per aprire l'ordine.")
        money_cfg = lambda label: st.column_config.NumberColumn(label, format="%.2f")
        sel = st.dataframe(
            g, hide_index=True, use_container_width=True,
            on_select="rerun", selection_mode="single-row", key="ov",
            column_config={"brand": "Brand", "stagione": "Stagione", "ordini": "Ordini",
                           "qty": "Qtà totale",
                           "importo": money_cfg("Importo totale (€)"),
                           "incassato": money_cfg("Incassato (€)"),
                           "da_incassare": money_cfg("Da incassare (€)"),
                           "da_saldare": "Pagamenti da saldare",
                           "scaduti": "Pagamenti scaduti",
                           "da_spedire": "Spedizioni da fare"})
        if sel.selection.rows:
            r = g.iloc[sel.selection.rows[0]]
            st.session_state["_open"] = (r["brand"], r["stagione"])
            st.rerun()

# ---------- Ordini ----------
if view == "Ordini":
    if "flash" in st.session_state:
        st.success(st.session_state.pop("flash"))
    if not (brand and season):
        st.info("Scegli brand e stagione nella barra laterale, oppure importa un topsheet nella scheda Importa.")
    else:
        mask = (orders["brand"] == brand) & (orders["stagione"] == season)
        full = orders[mask].reset_index(drop=True)
        late = full[(full["pay_status"] != "saldato") & full["pay_due"].notna()
                    & (full["pay_due"] < datetime.now().date())]
        _, _, resid = money(full)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Ordini", len(full))
        c2.metric("Pagamenti da saldare", int((full["pay_status"] != "saldato").sum()))
        c3.metric("Pagamenti scaduti", len(late))
        c4.metric("Da incassare", eur(resid.sum()))

        # filtri: mostrano solo una parte delle righe, le altre restano intatte al salvataggio
        f1, f2, f3 = st.columns([2, 1, 1])
        q = f1.text_input("Cerca (cliente, trade name, città)", key="f_q").strip().lower()
        pf = f2.multiselect("Pagamento", PAY, key="f_pay")
        sf = f3.multiselect("Spedizione", SHIP, key="f_ship")
        cur = full.copy()
        if q:
            hit = cur[["cliente", "trade_name", "city"]].astype(str).apply(
                lambda c: c.str.lower().str.contains(q, regex=False)).any(axis=1)
            cur = cur[hit]
        if pf:
            cur = cur[cur["pay_status"].isin(pf)]
        if sf:
            cur = cur[cur["ship_status"].isin(sf)]
        cur = cur.reset_index(drop=True)
        cur["saldato"] = cur["pay_status"] == "saldato"

        if len(cur) != len(full):
            st.caption(f"Mostro {len(cur)} ordini su {len(full)}.")
        st.caption("Modifica direttamente nella tabella, poi premi Salva. Puoi aggiungere righe in fondo. "
                   "Per segnare un cliente come saldato spunta la casella «Saldato».")
        cfg = {
            "cliente": st.column_config.SelectboxColumn("Cliente", options=clients, required=True),
            "trade_name": "Trade name", "city": "Città", "qty": "Qtà",
            "pct": st.column_config.NumberColumn("Importo (€)", format="%.2f", min_value=0),
            "sel": "SEL", "oc": "OC", "prod": "PROD", "payment_terms": "Termini di pagamento",
            "producer_code": f"{season}", "rc_del": "RC DEL", "mje_del": "MJE DEL", "comments": "Commenti",
            "saldato": st.column_config.CheckboxColumn("Saldato"),
            "pay_status": st.column_config.SelectboxColumn("Pagamento", options=PAY),
            "pay_paid": "Importo pagato",
            "pay_due": st.column_config.DateColumn("Scadenza pag."),
            "ship_status": st.column_config.SelectboxColumn("Spedizione", options=SHIP),
            "ship_date": st.column_config.DateColumn("Data spedizione"),
            "updated_at": st.column_config.TextColumn("Ultima modifica", disabled=True),
            "updated_by": st.column_config.TextColumn("Modificato da", disabled=True),
        }
        col_order = EDIT[:EDIT.index("pay_status")] + ["saldato", "pay_status"] + EDIT[EDIT.index("pay_status") + 1:]
        edkey = f"ed-{brand}-{season}-{q}-{'.'.join(pf)}-{'.'.join(sf)}-{st.session_state.get('edv', 0)}"
        ed = st.data_editor(cur, column_config=cfg, column_order=col_order + ["updated_at", "updated_by"],
                            num_rows="dynamic", hide_index=True, use_container_width=True, key=edkey)
        if st.button("Salva modifiche", type="primary"):
            if not who.strip():
                st.error("Scrivi il tuo nome nella barra laterale prima di salvare.")
            else:
                now, rows = datetime.now().strftime("%Y-%m-%d %H:%M"), []
                ed = ed.copy()
                ed["brand"], ed["stagione"] = brand, season
                ed["id"] = ed["id"].apply(lambda x: x if isinstance(x, str) and x else uuid.uuid4().hex[:8])
                # la casella «Saldato» e il menu «Pagamento» devono restare coerenti
                was = dict(zip(cur["id"], cur["saldato"]))

                def pay_of(r):
                    flag = bool(r["saldato"]) if pd.notna(r["saldato"]) else False
                    if flag != was.get(r["id"], False):
                        if flag:
                            return "saldato"
                        return "acconto_pagato" if pd.notna(r["pay_paid"]) and r["pay_paid"] > 0 else "da_pagare"
                    return r["pay_status"]

                ed["pay_status"] = ed.apply(pay_of, axis=1)
                ed["pay_status"] = ed["pay_status"].where(ed["pay_status"].isin(PAY), PAY[0])
                ed["ship_status"] = ed["ship_status"].where(ed["ship_status"].isin(SHIP), SHIP[0])

                old, new = cur.set_index("id"), ed.set_index("id")
                added = [i for i in new.index if i not in old.index]
                changed = {}
                for i in new.index:
                    if i in old.index:
                        ch = [c for c in EDIT if not same(new.loc[i, c], old.loc[i, c])]
                        if ch:
                            changed[i] = ch
                deleted = list(old.index.difference(new.index))

                # dati freschi dal foglio: se qualcun altro ha toccato le stesse righe, non si salva
                fresh = load_orders(fresh=True)
                fcur = fresh.set_index("id")
                conflicts = [i for i in list(changed) + deleted
                             if i not in fcur.index
                             or norm(fcur.loc[i, "updated_at"]) != norm(old.loc[i, "updated_at"])]
                if conflicts:
                    nomi = ", ".join(sorted({norm(old.loc[i, "trade_name"]) for i in conflicts}))
                    st.error(f"Queste righe sono state modificate da qualcun altro mentre lavoravi: {nomi}. "
                             "Non è stato salvato nulla: premi «Ricarica dati» e rifai la modifica.")
                elif not (added or changed or deleted):
                    st.info("Nessuna modifica da salvare.")
                else:
                    for i in added:
                        new.loc[i, ["updated_at", "updated_by"]] = [now, who]
                        rows.append([now, who, i, new.loc[i, "trade_name"], "(nuova riga)", "", ""])
                    for i, ch in changed.items():
                        new.loc[i, ["updated_at", "updated_by"]] = [now, who]
                        rows += [[now, who, i, new.loc[i, "trade_name"], c, norm(old.loc[i, c]), norm(new.loc[i, c])]
                                 for c in ch]
                    rows += [[now, who, i, old.loc[i, "trade_name"], "(riga eliminata)", "", ""] for i in deleted]
                    # si parte dai dati freschi e si applicano solo le righe toccate da te
                    touched = set(changed) | set(added) | set(deleted)
                    base = fcur[~fcur.index.isin(touched)]
                    mine = new[new.index.isin(set(changed) | set(added))]
                    keep = [c for c in COLS if c != "id"]
                    merged = pd.concat([base[keep], mine[keep]])
                    merged = merged.loc[[i for i in list(fresh["id"]) + added if i in merged.index]]
                    save_orders(merged.reset_index()[COLS])
                    log(rows)
                    st.session_state["edv"] = st.session_state.get("edv", 0) + 1
                    st.session_state["flash"] = f"Salvato: {len(rows)} modifiche registrate."
                    st.rerun()

# ---------- Importa ----------
if view == "Importa":
    up = st.file_uploader("Carica un topsheet (.xlsx)", type=["xlsx"])
    if up:
        try:
            b, s, df = parse_xlsx(up)
        except Exception as e:
            st.error(f"Importazione non riuscita: {e}")
        else:
            b = st.text_input("Brand", b)
            s = st.text_input("Stagione", s)
            key = lambda d: (d["brand"] + "|" + d["stagione"] + "|" + d["trade_name"].str.lower() + "|" + d["city"].str.lower())
            df = df.assign(brand=b, stagione=s)
            n_dup = int(key(df).isin(set(key(orders))).sum()) if len(orders) else 0
            st.dataframe(df.drop(columns=["brand", "stagione"]), hide_index=True, use_container_width=True)
            upd = st.checkbox("Aggiorna i dati degli ordini già presenti "
                              "(stato pagamento, importo pagato e spedizione non vengono toccati)", value=True)
            st.caption(f"{len(df) - n_dup} ordini nuovi, {n_dup} già presenti"
                       + (" (verranno aggiornati)." if upd and n_dup else " (verranno saltati)." if n_dup else "."))
            if st.button(f"Importa {len(df)} ordini in {b} {s}", type="primary"):
                if not who.strip():
                    st.error("Scrivi il tuo nome nella barra laterale.")
                else:
                    now = datetime.now().strftime("%Y-%m-%d %H:%M")
                    fresh = load_orders(fresh=True)
                    fkeys = key(fresh) if len(fresh) else pd.Series([], dtype=object)
                    dup_mask = key(df).isin(set(fkeys))
                    dup, df = df[dup_mask], df[~dup_mask]
                    rows_log, n_upd = [], 0
                    if upd and len(dup):
                        pos = dict(zip(fkeys, fresh.index))
                        for r in dup.assign(_k=key(dup)).to_dict("records"):
                            i = pos[r["_k"]]
                            ch = [c for c in UPD if not same(fresh.at[i, c], r[c])]
                            if ch:
                                n_upd += 1
                                rows_log += [[now, who, fresh.at[i, "id"], fresh.at[i, "trade_name"], c,
                                              norm(fresh.at[i, c]), norm(r[c])] for c in ch]
                                for c in ch:
                                    fresh.at[i, c] = r[c]
                                fresh.at[i, "updated_at"], fresh.at[i, "updated_by"] = now, who
                    if len(df):
                        df = df.copy()
                        df["id"] = [uuid.uuid4().hex[:8] for _ in range(len(df))]
                        df["updated_at"], df["updated_by"] = now, who
                        for c in COLS:
                            if c not in df:
                                df[c] = None
                        rows_log += [[now, who, r.id, r.trade_name, "(importato)", "", ""] for r in df.itertuples()]
                        fresh = pd.concat([fresh, df[COLS]], ignore_index=True)
                    if len(df) or n_upd:
                        save_orders(fresh[COLS])
                        log(rows_log)
                    if b not in brands:
                        write("Brand", pd.DataFrame({"name": brands + [b]}), ANA["Brand"])
                    if not ((seasons["brand"] == b) & (seasons["code"] == s)).any():
                        write("Stagioni", pd.concat([seasons, pd.DataFrame([{"brand": b, "code": s}])]), ANA["Stagioni"])
                    new_c = [c for c in df["cliente"].unique() if c not in clients] if len(df) else []
                    if new_c:
                        write("Clienti", pd.DataFrame({"name": clients + new_c}), ANA["Clienti"])
                    st.success(f"Importati {len(df)} ordini nuovi, aggiornati {n_upd} già presenti"
                               f"{'' if upd else ' (i duplicati sono stati saltati)'}.")

# ---------- Anagrafiche ----------
if view == "Anagrafiche":
    st.caption("Elenchi usati nei menu. Aggiungi righe in fondo e premi Salva.")
    for name, cols in ANA.items():
        st.subheader(name)
        data = read(name, tuple(cols))
        out = st.data_editor(data, num_rows="dynamic", hide_index=True, key=f"ana-{name}")
        if st.button(f"Salva {name.lower()}", key=f"s-{name}"):
            write(name, out.dropna(how="all"), cols)
            st.rerun()

# ---------- PDF ----------
if view == "PDF":
    if brand and season:
        sub = orders[(orders["brand"] == brand) & (orders["stagione"] == season)]
        st.write(f"Topsheet **{brand} {season}**, {len(sub)} ordini, con i dati salvati finora.")
        st.download_button("Scarica PDF", make_pdf(sub, brand, season),
                           file_name=f"{brand}_{season}_TOPSHEET_{datetime.now():%Y%m%d}.pdf".replace(" ", "_"),
                           mime="application/pdf", type="primary")
    else:
        st.info("Scegli brand e stagione nella barra laterale.")

# ---------- Log ----------
if view == "Log":
    lg = read("Log", tuple(LOGCOLS))
    if lg.empty:
        st.info("Nessuna modifica registrata finora.")
    else:
        l1, l2 = st.columns([2, 1])
        t = l1.text_input("Cerca (trade name, campo, valore)", key="log_q").strip().lower()
        chi = l2.selectbox("Autore", ["Tutti"] + sorted(x for x in lg["chi"].unique() if x), key="log_chi")
        if t:
            hit = lg[["trade_name", "campo", "prima", "dopo"]].astype(str).apply(
                lambda c: c.str.lower().str.contains(t, regex=False)).any(axis=1)
            lg = lg[hit]
        if chi != "Tutti":
            lg = lg[lg["chi"] == chi]
        lg = lg.sort_values("quando", ascending=False)
        st.caption(f"{len(lg)} modifiche (mostro le ultime 1000).")
        st.dataframe(lg.head(1000).drop(columns=["id"]), hide_index=True, use_container_width=True,
                     column_config={"quando": "Quando", "chi": "Chi", "trade_name": "Trade name",
                                    "campo": "Campo", "prima": "Prima", "dopo": "Dopo"})
