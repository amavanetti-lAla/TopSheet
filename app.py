import json
import re
import time
import uuid
from datetime import datetime

import gspread
import pandas as pd
import streamlit as st

from core import COLS, PAY, SHIP, make_pdf, num, parse_xlsx

st.set_page_config(page_title="Topsheet", layout="wide")

ANA = {"Brand": ["name"], "Stagioni": ["brand", "code"], "Clienti": ["name"]}
LOGCOLS = ["quando", "chi", "id", "trade_name", "campo", "prima", "dopo"]
EDIT = ["cliente", "trade_name", "city", "qty", "pct", "sel", "oc", "prod", "payment_terms", "producer_code",
        "rc_del", "mje_del", "comments", "pay_status", "pay_paid", "pay_due", "ship_status", "ship_date"]
VIEWS = ["Panoramica", "Ordini", "Importa", "Anagrafiche", "PDF"]


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


@st.cache_data(ttl=120, show_spinner=False)
def read(name, cols):
    w = sheet(name, tuple(cols))
    rec = with_retry(lambda: w.get_all_records(numericise_ignore=["all"], default_blank=""))
    df = pd.DataFrame(rec)
    for c in cols:
        if c not in df:
            df[c] = ""
    return df[list(cols)]


def write(name, df, cols):
    w = sheet(name, tuple(cols))
    values = [list(cols)] + df[list(cols)].fillna("").astype(str).values.tolist()
    with_retry(w.clear)
    with_retry(lambda: w.update(values, value_input_option="RAW"))
    st.cache_data.clear()  # le prossime letture prendono i dati aggiornati


def log(rows):
    if rows:
        w = sheet("Log", tuple(LOGCOLS))
        with_retry(lambda: w.append_rows(rows, value_input_option="RAW"))


def load_orders():
    df = read("Ordini", tuple(COLS))
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


def season_key(code):
    """SS26 < FW26 < SS27 ...: ordina per anno, poi primavera/estate prima di autunno/inverno."""
    c = str(code).upper().strip()
    m = re.search(r"(\d{2})\s*$", c)
    year = int(m.group(1)) if m else 0
    half = 0 if c.startswith("PS") else 1 if c.startswith(("SS", "S", "PE", "RE")) else 2
    return (year, half)


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
        st.rerun()

view = st.radio("Vista", VIEWS, horizontal=True, key="view", label_visibility="collapsed")

# ---------- Panoramica ----------
if view == "Panoramica":
    if orders.empty:
        st.info("Nessun topsheet caricato: usa la scheda Importa.")
    else:
        recenti = st.toggle("Dalla stagione più recente", value=True)
        g = (orders.assign(da_saldare=orders["pay_status"] != "saldato")
             .groupby(["brand", "stagione"])
             .agg(ordini=("id", "count"), qty=("qty", "sum"), importo=("pct", "sum"),
                  da_saldare=("da_saldare", "sum"))
             .reset_index())
        g["_k"] = g["stagione"].map(season_key)
        g = (g.sort_values(["_k", "brand"], ascending=[not recenti, True])
              .drop(columns="_k").reset_index(drop=True))
        st.caption("Clicca su una riga per aprire l'ordine.")
        sel = st.dataframe(
            g, hide_index=True, use_container_width=True,
            on_select="rerun", selection_mode="single-row", key="ov",
            column_config={"brand": "Brand", "stagione": "Stagione", "ordini": "Ordini",
                           "qty": "Qtà totale",
                           "importo": st.column_config.NumberColumn("Importo totale (€)", format="%.2f"),
                           "da_saldare": "Pagamenti da saldare"})
        if sel.selection.rows:
            r = g.iloc[sel.selection.rows[0]]
            st.session_state["_open"] = (r["brand"], r["stagione"])
            st.rerun()

# ---------- Ordini ----------
if view == "Ordini":
    if not (brand and season):
        st.info("Scegli brand e stagione nella barra laterale, oppure importa un topsheet nella scheda Importa.")
    else:
        mask = (orders["brand"] == brand) & (orders["stagione"] == season)
        cur = orders[mask].reset_index(drop=True)
        late = cur[(cur["pay_status"] != "saldato") & cur["pay_due"].notna() & (cur["pay_due"] < datetime.now().date())]
        c1, c2, c3 = st.columns(3)
        c1.metric("Ordini", len(cur))
        c2.metric("Pagamenti da saldare", int((cur["pay_status"] != "saldato").sum()))
        c3.metric("Pagamenti scaduti", len(late))
        st.caption("Modifica direttamente nella tabella, poi premi Salva. Puoi aggiungere righe in fondo.")
        cfg = {
            "cliente": st.column_config.SelectboxColumn("Cliente", options=clients, required=True),
            "trade_name": "Trade name", "city": "Città", "qty": "Qtà",
            "pct": st.column_config.NumberColumn("Importo (€)", format="%.2f", min_value=0),
            "sel": "SEL", "oc": "OC", "prod": "PROD", "payment_terms": "Termini di pagamento",
            "producer_code": f"{season}", "rc_del": "RC DEL", "mje_del": "MJE DEL", "comments": "Commenti",
            "pay_status": st.column_config.SelectboxColumn("Pagamento", options=PAY),
            "pay_paid": "Importo pagato",
            "pay_due": st.column_config.DateColumn("Scadenza pag."),
            "ship_status": st.column_config.SelectboxColumn("Spedizione", options=SHIP),
            "ship_date": st.column_config.DateColumn("Data spedizione"),
            "updated_at": st.column_config.TextColumn("Ultima modifica", disabled=True),
            "updated_by": st.column_config.TextColumn("Modificato da", disabled=True),
        }
        ed = st.data_editor(cur, column_config=cfg, column_order=EDIT + ["updated_at", "updated_by"],
                            num_rows="dynamic", hide_index=True, use_container_width=True, key=f"ed-{brand}-{season}")
        if st.button("Salva modifiche", type="primary"):
            if not who.strip():
                st.error("Scrivi il tuo nome nella barra laterale prima di salvare.")
            else:
                now, rows = datetime.now().strftime("%Y-%m-%d %H:%M"), []
                ed = ed.copy()
                ed["brand"], ed["stagione"] = brand, season
                ed["id"] = ed["id"].apply(lambda x: x if isinstance(x, str) and x else uuid.uuid4().hex[:8])
                old, new = cur.set_index("id"), ed.set_index("id")
                for i in new.index:
                    if i not in old.index:
                        new.loc[i, ["updated_at", "updated_by"]] = [now, who]
                        rows.append([now, who, i, new.loc[i, "trade_name"], "(nuova riga)", "", ""])
                        continue
                    ch = [c for c in EDIT if norm(new.loc[i, c]) != norm(old.loc[i, c])]
                    if ch:
                        new.loc[i, ["updated_at", "updated_by"]] = [now, who]
                    rows += [[now, who, i, new.loc[i, "trade_name"], c, norm(old.loc[i, c]), norm(new.loc[i, c])] for c in ch]
                rows += [[now, who, i, old.loc[i, "trade_name"], "(riga eliminata)", "", ""] for i in old.index.difference(new.index)]
                save_orders(pd.concat([orders[~mask], new.reset_index()], ignore_index=True))
                log(rows)
                st.success(f"Salvato: {len(rows)} modifiche registrate.")
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
            st.dataframe(df, hide_index=True, use_container_width=True)
            if st.button(f"Importa {len(df)} ordini in {b} {s}", type="primary"):
                if not who.strip():
                    st.error("Scrivi il tuo nome nella barra laterale.")
                else:
                    now = datetime.now().strftime("%Y-%m-%d %H:%M")
                    key = lambda d: (d["brand"] + "|" + d["stagione"] + "|" + d["trade_name"].str.lower() + "|" + d["city"].str.lower())
                    df = df.assign(brand=b, stagione=s)
                    df = df[~key(df).isin(set(key(orders)))]
                    if len(df):
                        df["id"] = [uuid.uuid4().hex[:8] for _ in range(len(df))]
                        df["updated_at"], df["updated_by"] = now, who
                        for c in COLS:
                            if c not in df:
                                df[c] = None
                        save_orders(pd.concat([orders, df[COLS]], ignore_index=True))
                        log([[now, who, r.id, r.trade_name, "(importato)", "", ""] for r in df.itertuples()])
                    if b not in brands:
                        write("Brand", pd.DataFrame({"name": brands + [b]}), ANA["Brand"])
                    if not ((seasons["brand"] == b) & (seasons["code"] == s)).any():
                        write("Stagioni", pd.concat([seasons, pd.DataFrame([{"brand": b, "code": s}])]), ANA["Stagioni"])
                    new_c = [c for c in df["cliente"].unique() if c not in clients] if len(df) else []
                    if new_c:
                        write("Clienti", pd.DataFrame({"name": clients + new_c}), ANA["Clienti"])
                    st.success(f"Importati {len(df)} ordini (i duplicati sono stati saltati).")

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
