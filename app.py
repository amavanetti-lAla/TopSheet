import json
import re
import time
import uuid
from datetime import datetime

import gspread
import pandas as pd
import streamlit as st

from core import COLS, CURRENCIES, PAY, SHIP, cur_sym, eur, make_pdf, num
from topsheet_parser import parse_xlsx

st.set_page_config(page_title="Topsheet", layout="wide")

WSP_URL = "https://wsppriceselectiontool.streamlit.app/"

ANA = {"Brand": ["name"], "Stagioni": ["brand", "code"], "Clienti": ["name"]}
# anagrafica clienti e persone di riferimento (schede separate: la scheda "Clienti" resta l'elenco dei nomi)
CLI_COLS = ["name", "azienda", "attivo", "indirizzo", "citta", "paese", "partita_iva", "note"]
CON_COLS = ["id", "cliente", "nome", "cognome", "ruolo", "email", "telefono", "attivo", "note"]
LOGCOLS = ["quando", "chi", "id", "trade_name", "campo", "prima", "dopo"]
EDIT = ["cliente", "trade_name", "city", "qty", "pct", "valuta", "sel", "oc", "prod", "payment_terms", "producer_code",
        "rc_del", "mje_del", "comments", "pay_status", "pay_paid", "pay_due", "ship_status", "ship_date"]
# campi che un reimport può aggiornare (stato pagamento e spedizione non vengono mai toccati)
UPD = ["cliente", "qty", "pct", "valuta", "sel", "oc", "prod", "payment_terms", "producer_code", "rc_del", "mje_del",
       "comments"]
VIEWS = ["Panoramica", "Ordini", "Anagrafiche", "Importa", "PDF", "Log"]
# etichette con semaforo mostrate nei menu Pagamento / Spedizione (nel foglio restano i codici PAY / SHIP)
PAY_LBL = {"da_pagare": "🔴 Da pagare", "acconto_pagato": "🟡 Acconto pagato", "saldato": "🟢 Saldato"}
SHIP_LBL = {"da_spedire": "🔴 Da spedire", "spedito": "🟡 Spedito", "consegnato": "🟢 Consegnato"}
PAY_REV = {v: k for k, v in PAY_LBL.items()}
SHIP_REV = {v: k for k, v in SHIP_LBL.items()}


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
    with_retry(lambda: w.resize(rows=len(values) + 100, cols=len(cols)))  # la scheda cresce (o si riduce) secondo i dati
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
    df["valuta"] = df["valuta"].map(cur_sym)
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


def money_metric(col, label, frame, colname):
    """Come st.metric, ma con una sola valuta mostra il valore normale; con più valute le impila una sopra l'altra
    (mai somme miste € e $)."""
    tot = frame.groupby("valuta")[colname].sum()
    items = sorted(tot.items(), key=lambda kv: kv[0] != "€")
    if len(items) <= 1:
        col.metric(label, eur(items[0][1], items[0][0]) if items else eur(0))
    else:
        lines = "".join(f'<div style="font-size:1.6rem;line-height:1.35">{eur(v, c)}</div>' for c, v in items)
        col.markdown(f'<div style="font-size:14px;opacity:.7">{label}</div>{lines}', unsafe_allow_html=True)


# ---------- Dati ----------
orders = load_orders()
brands = read("Brand", tuple(ANA["Brand"]))["name"].tolist()
seasons = read("Stagioni", tuple(ANA["Stagioni"]))


def pinned_col(factory, *args, **kw):
    """Colonna fissata a sinistra mentre si scorre; con versioni di Streamlit più vecchie è una colonna normale."""
    try:
        return factory(*args, pinned=True, **kw)
    except TypeError:
        return factory(*args, **kw)


def cur_text(frame, col):
    """Somma di una colonna scritta per valuta: '€776,00 + $48.244,40' (mai somme miste)."""
    tot = frame.groupby("valuta")[col].sum()
    parts = [eur(v, c) for c, v in sorted(tot.items(), key=lambda kv: kv[0] != "€")]
    return " + ".join(parts) or "—"


def sort_names(names):
    """Elenco senza vuoti né doppioni, in ordine alfabetico (maiuscole/minuscole ininfluenti)."""
    return sorted({str(n).strip() for n in names if str(n).strip()}, key=str.casefold)


clients = sort_names(read("Clienti", tuple(ANA["Clienti"]))["name"].tolist())

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

# Barra superiore: selettore di vista a sinistra, link per tornare a WSP a destra
view_col, link_col = st.columns([4, 1])
with view_col:
    view = st.radio("Vista", VIEWS, horizontal=True, key="view", label_visibility="collapsed")
with link_col:
    st.link_button("Apri WSP ↗", WSP_URL, use_container_width=True)

# ---------- Panoramica ----------
if view == "Panoramica":
    if "flash_pan" in st.session_state:
        st.success(st.session_state.pop("flash_pan"))
    if orders.empty:
        st.info("Nessun topsheet caricato: usa la scheda Importa.")
    else:
        o = orders.copy()
        o["importo"], o["incassato"], o["da_incassare"] = money(o)
        saldato = o["pay_status"] == "saldato"
        o["da_saldare"] = ~saldato
        o["scaduti"] = ~saldato & o["pay_due"].notna() & (o["pay_due"] < datetime.now().date())
        o["da_spedire"] = o["ship_status"] == "da_spedire"
        # una riga per brand e stagione; gli importi sono scritti per valuta ("€776,00 + $48.244,40")
        g = (o.groupby(["brand", "stagione"])
             .agg(ordini=("id", "count"), qty=("qty", "sum"),
                  da_saldare=("da_saldare", "sum"), scaduti=("scaduti", "sum"),
                  da_spedire=("da_spedire", "sum"))
             .reset_index())
        for col in ("importo", "incassato", "da_incassare"):
            parts = {}
            tot = o.groupby(["brand", "stagione", "valuta"])[col].sum()
            for (b_, s_, c_), v_ in sorted(tot.items(), key=lambda kv: kv[0][2] != "€"):
                parts.setdefault((b_, s_), []).append(eur(v_, c_))
            g[col] = [" + ".join(parts.get((b_, s_), [])) for b_, s_ in zip(g["brand"], g["stagione"])]
        g = g[["brand", "stagione", "ordini", "qty", "importo", "incassato", "da_incassare",
               "da_saldare", "scaduti", "da_spedire"]]
        n_scad = int(g["scaduti"].sum())
        if n_scad:  # nessun totale generale (mescolerebbe stagioni e valute): solo un avviso se ci sono scadenze
            st.warning(f"⚠️ {n_scad} pagamenti scaduti: guarda la colonna «Pagamenti scaduti» qui sotto.")
        recenti = st.toggle("Dalla stagione più recente", value=True)
        g["_k"] = g["stagione"].map(season_key)
        g = (g.sort_values(["_k", "brand"], ascending=[not recenti, True])
              .drop(columns="_k").reset_index(drop=True))
        st.caption("Clicca su una riga per aprire l'ordine.")
        sel = st.dataframe(
            g, hide_index=True, use_container_width=True,
            on_select="rerun", selection_mode="single-row", key="ov",
            column_config={"brand": "Brand", "stagione": "Stagione", "ordini": "Ordini",
                           "qty": "Qtà totale",
                           "importo": "Importo totale",
                           "incassato": "Incassato",
                           "da_incassare": "Da incassare",
                           "da_saldare": "Pagamenti da saldare",
                           "scaduti": "Pagamenti scaduti",
                           "da_spedire": "Spedizioni da fare"})
        if sel.selection.rows:
            r = g.iloc[sel.selection.rows[0]]
            st.session_state["_open"] = (r["brand"], r["stagione"])
            st.rerun()

        with st.expander("🗑️ Elimina un intero topsheet (tutti gli ordini di un brand e stagione)"):
            pv = st.session_state.get("pdv", 0)  # cambia dopo ogni eliminazione: azzera menu e conferma
            combos = list(g[["brand", "stagione"]].drop_duplicates().itertuples(index=False, name=None))
            pick = st.selectbox("Topsheet da eliminare", combos, key=f"pan_del_pick_{pv}",
                                format_func=lambda t: f"{t[0]} {t[1]}")
            n_ord = int(((orders["brand"] == pick[0]) & (orders["stagione"] == pick[1])).sum())
            st.warning(f"Verranno eliminati {n_ord} ordini di {pick[0]} {pick[1]} (in tutte le valute). "
                       "Prima dell'eliminazione ne resta una copia nel foglio «Backup_Ordini» del Google Sheet.")
            ok = st.checkbox("Confermo: voglio eliminare questi ordini", key=f"pan_del_ok_{pv}")
            if st.button("Elimina topsheet", type="primary", disabled=not ok, key=f"pan_del_btn_{pv}"):
                if not who.strip():
                    st.error("Scrivi il tuo nome nella barra laterale prima di eliminare.")
                else:
                    fresh = load_orders(fresh=True)
                    gone = fresh[(fresh["brand"] == pick[0]) & (fresh["stagione"] == pick[1])]
                    if gone.empty:
                        st.info("Questo topsheet non ha più ordini: premi «Ricarica dati».")
                    else:
                        now = datetime.now().strftime("%Y-%m-%d %H:%M")
                        rows = [[now, who, r.id, r.trade_name, "(riga eliminata)", "", ""] for r in gone.itertuples()]
                        save_orders(fresh.drop(gone.index)[COLS])
                        log(rows)
                        st.session_state["pdv"] = pv + 1
                        st.session_state["flash_pan"] = f"Eliminati {len(gone)} ordini di {pick[0]} {pick[1]}."
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
        money_metric(c4, "Da incassare", full.assign(_r=resid), "_r")

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
        cur["elimina"] = False

        # semafori (colonne di sola lettura, si aggiornano dopo il Salva)
        today = datetime.now().date()

        def sem_pay(r):
            if r["pay_status"] == "saldato":
                return "🟢 Saldato"
            d = r["pay_due"]
            if d is not None and pd.notna(d):
                if d < today:
                    return "🔴 Scaduto"
                if (d - today).days <= 7:
                    return "🟠 In scadenza"
            return "🟡 Da saldare"

        cur["sem_pay"] = [sem_pay(r) for _, r in cur.iterrows()]
        cur_codes = cur.copy()  # versione con i codici del foglio, per confrontare le modifiche al Salva
        cur["pay_status"] = cur["pay_status"].map(PAY_LBL)
        cur["ship_status"] = cur["ship_status"].map(SHIP_LBL)

        if len(cur) != len(full):
            st.caption(f"Mostro {len(cur)} ordini su {len(full)}.")
        st.caption("Modifica direttamente nella tabella, poi premi Salva. Puoi aggiungere righe in fondo. "
                   "Per segnare un cliente come saldato spunta la casella «Saldato». "
                   "Per eliminare un ordine spunta «Elimina» (la riga sparisce al Salva). "
                   "I semafori 🟢🟡🟠🔴 della colonna «Scadenza» si aggiornano dopo il Salva.")
        cli_opts = sorted(set(clients) | {str(c) for c in full["cliente"] if str(c).strip()}, key=str.casefold)
        # larghezze ridotte e intestazioni brevi, così la tabella sta in una schermata
        cfg = {
            "elimina": st.column_config.CheckboxColumn("🗑️", width="small", help="Spunta per eliminare l'ordine al Salva"),
            "cliente": pinned_col(st.column_config.SelectboxColumn, "Cliente", options=cli_opts, required=True),
            "trade_name": pinned_col(st.column_config.TextColumn, "Trade name"),
            "city": st.column_config.TextColumn("Città"),
            "qty": st.column_config.NumberColumn("Qtà", width="small"),
            "pct": st.column_config.NumberColumn("Importo", format="%.2f", min_value=0),
            "valuta": st.column_config.SelectboxColumn("Val.", options=CURRENCIES, required=True, width="small"),
            "sel": st.column_config.CheckboxColumn("SEL", width="small"),
            "oc": st.column_config.CheckboxColumn("OC", width="small"),
            "prod": st.column_config.CheckboxColumn("PROD", width="small"),
            "payment_terms": st.column_config.TextColumn("Termini"),
            "producer_code": st.column_config.TextColumn(f"{season}", width="small"),
            "rc_del": st.column_config.TextColumn("RC DEL"),
            "mje_del": st.column_config.TextColumn("MJE DEL"),
            "comments": st.column_config.TextColumn("Commenti"),
            "sem_pay": st.column_config.TextColumn("Scadenza", disabled=True),
            "saldato": st.column_config.CheckboxColumn("Saldato", width="small"),
            "pay_status": st.column_config.SelectboxColumn("Pagamento", options=list(PAY_LBL.values())),
            "pay_paid": st.column_config.NumberColumn("Pagato", width="small"),
            "pay_due": st.column_config.DateColumn("Scad. pag.", format="DD/MM/YYYY"),
            "ship_status": st.column_config.SelectboxColumn("Spedizione", options=list(SHIP_LBL.values())),
            "ship_date": st.column_config.DateColumn("Data sped.", format="DD/MM/YYYY"),
            "updated_at": st.column_config.TextColumn("Ultima modifica", disabled=True),
            "updated_by": st.column_config.TextColumn("Modificato da", disabled=True),
        }
        col_order = EDIT[:EDIT.index("pay_status")] + ["saldato", "pay_status"] + EDIT[EDIT.index("pay_status") + 1:]
        # il semaforo di scadenza sta davanti a «Saldato»
        col_order = [x for c in col_order for x in (["sem_pay", c] if c == "saldato" else [c])]
        edkey = f"ed-{brand}-{season}-{q}-{'.'.join(pf)}-{'.'.join(sf)}-{st.session_state.get('edv', 0)}"
        show_audit = st.toggle("Mostra anche «Ultima modifica» e «Modificato da»", key="show_audit")
        cols_shown = col_order + (["updated_at", "updated_by"] if show_audit else []) + ["elimina"]
        # altezza = tutte le righe visibili senza scorrere (fino a un massimo)
        ed = st.data_editor(cur, column_config=cfg, column_order=cols_shown,
                            num_rows="dynamic", hide_index=True, use_container_width=True, key=edkey,
                            height=min(35 * (len(cur) + 2) + 3, 1000))
        n_del = int(ed["elimina"].fillna(False).astype(bool).sum())
        if n_del:
            st.warning(f"{n_del} ordini con «Elimina» spuntato verranno eliminati quando premi Salva.")
        if st.button("Salva modifiche", type="primary"):
            if not who.strip():
                st.error("Scrivi il tuo nome nella barra laterale prima di salvare.")
            else:
                now, rows = datetime.now().strftime("%Y-%m-%d %H:%M"), []
                ed = ed.copy()
                ed["brand"], ed["stagione"] = brand, season
                ed["valuta"] = ed["valuta"].where(ed["valuta"].isin(CURRENCIES), "€")  # righe nuove: euro
                ed["id"] = ed["id"].apply(lambda x: x if isinstance(x, str) and x else uuid.uuid4().hex[:8])
                # righe con «Elimina» spuntato: tolte dalla tabella, quindi finiscono tra le righe eliminate
                ed = ed[~ed["elimina"].fillna(False).astype(bool)]
                ed["pay_status"] = ed["pay_status"].map(PAY_REV)    # dalle etichette con semaforo ai codici
                ed["ship_status"] = ed["ship_status"].map(SHIP_REV)
                # la casella «Saldato» e il menu «Pagamento» devono restare coerenti
                was = dict(zip(cur["id"], cur["saldato"]))

                def pay_of(r):
                    flag = bool(r["saldato"]) if pd.notna(r["saldato"]) else False
                    if flag != was.get(r["id"], False):
                        if flag:
                            return "saldato"
                        return "acconto_pagato" if pd.notna(r["pay_paid"]) and r["pay_paid"] > 0 else "da_pagare"
                    return r["pay_status"]

                ed["pay_status"] = [pay_of(r) for _, r in ed.iterrows()]  # (funziona anche se non resta nessuna riga)
                ed["pay_status"] = ed["pay_status"].where(ed["pay_status"].isin(PAY), PAY[0])
                ed["ship_status"] = ed["ship_status"].where(ed["ship_status"].isin(SHIP), SHIP[0])

                old, new = cur_codes.set_index("id"), ed.set_index("id")
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
            canon = {" ".join(c.split()).casefold(): c for c in clients}  # stesso cliente scritto in modo diverso
            df["cliente"] = df["cliente"].map(
                lambda c: canon.get(" ".join(str(c).split()).casefold(), " ".join(str(c).split())))
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
                        write("Clienti", pd.DataFrame({"name": sort_names(clients + new_c)}), ANA["Clienti"])
                    st.success(f"Importati {len(df)} ordini nuovi, aggiornati {n_upd} già presenti"
                               f"{'' if upd else ' (i duplicati sono stati saltati)'}.")


# ---------- Scheda cliente (anagrafica + persone di riferimento) ----------
def scheda_cliente(c_sel, who):
    """Scheda di un cliente: anagrafica + persone di riferimento (con log delle modifiche)."""
    ana = read("Anagrafica_Clienti", tuple(CLI_COLS))
    cur = ana[ana["name"] == c_sel]
    a = cur.iloc[0].to_dict() if len(cur) else {c: "" for c in CLI_COLS}
    st.markdown(f"#### {c_sel}")

    with st.form(f"anag_form_{c_sel}"):
        c1, c2, c3 = st.columns(3)
        v = {}
        v["azienda"] = c1.text_input("Azienda", a["azienda"], key=f"an_az_{c_sel}")
        v["attivo"] = c2.selectbox("Attivo", ["SI", "NO"], index=1 if str(a["attivo"]).upper() == "NO" else 0,
                                   key=f"an_at_{c_sel}")
        v["partita_iva"] = c3.text_input("Partita IVA", a["partita_iva"], key=f"an_iva_{c_sel}")
        v["indirizzo"] = st.text_input("Indirizzo", a["indirizzo"], key=f"an_ind_{c_sel}")
        d1, d2 = st.columns(2)
        v["citta"] = d1.text_input("Città", a["citta"], key=f"an_ci_{c_sel}")
        v["paese"] = d2.text_input("Paese", a["paese"], key=f"an_pa_{c_sel}")
        v["note"] = st.text_area("Note", a["note"], key=f"an_no_{c_sel}")
        ok = st.form_submit_button("Salva scheda")
    if ok:
        if not who.strip():
            st.error("Scrivi il tuo nome nella barra laterale prima di salvare.")
        else:
            now = datetime.now().strftime("%Y-%m-%d %H:%M")
            fresh = _read("Anagrafica_Clienti", tuple(CLI_COLS))
            o = fresh[fresh["name"] == c_sel]
            old = o.iloc[0].to_dict() if len(o) else {c: "" for c in CLI_COLS}
            rows = [[now, who, "anagrafica", c_sel, k, norm(old[k]), norm(val)]
                    for k, val in v.items() if not same(old[k], val)]
            if rows:
                fresh = pd.concat([fresh[fresh["name"] != c_sel], pd.DataFrame([{"name": c_sel, **v}])],
                                  ignore_index=True)
                write("Anagrafica_Clienti", fresh.sort_values("name", key=lambda s: s.str.casefold()), CLI_COLS)
                log(rows)
                st.session_state["flash_ana"] = f"Scheda di {c_sel} salvata."
                st.rerun()
            else:
                st.info("Nessuna modifica da salvare.")

    st.markdown("**Persone di riferimento**")
    con = read("Contatti", tuple(CON_COLS))
    mine = con[con["cliente"] == c_sel].reset_index(drop=True)
    mine["attivo"] = mine["attivo"].where(mine["attivo"].isin(["SI", "NO"]), "SI")
    only_act = st.toggle("Solo attivi", key=f"con_act_{c_sel}")
    shown = mine[mine["attivo"] == "SI"].reset_index(drop=True) if only_act else mine
    ed = st.data_editor(
        shown, num_rows="dynamic", hide_index=True, use_container_width=True,
        column_order=["nome", "cognome", "ruolo", "email", "telefono", "attivo", "note"],
        column_config={"nome": "Nome", "cognome": "Cognome", "ruolo": "Ruolo", "email": "Email",
                       "telefono": "Telefono", "note": "Note",
                       "attivo": st.column_config.SelectboxColumn("Attivo", options=["SI", "NO"], width="small")},
        key=f"con_ed_{c_sel}_{only_act}_{st.session_state.get('conv', 0)}")
    if st.button("Salva contatti", key=f"con_save_{c_sel}"):
        if not who.strip():
            st.error("Scrivi il tuo nome nella barra laterale prima di salvare.")
        else:
            now = datetime.now().strftime("%Y-%m-%d %H:%M")
            e = ed.copy()
            e["cliente"] = c_sel
            e["attivo"] = e["attivo"].where(e["attivo"].isin(["SI", "NO"]), "SI")
            e["id"] = e["id"].apply(lambda x: x if isinstance(x, str) and x else uuid.uuid4().hex[:8])
            e = e.fillna("")
            vuoto = (e[["nome", "cognome", "email"]].astype(str).apply(lambda c: c.str.strip()) == "").all(axis=1)
            e = e[~vuoto][CON_COLS]
            fresh = _read("Contatti", tuple(CON_COLS))
            old = fresh[fresh["cliente"] == c_sel].set_index("id")
            new = e.set_index("id")
            if only_act:  # i contatti non attivi nascosti dal filtro restano com'erano
                hidden = old[~old.index.isin(shown["id"])]
                new = pd.concat([new, hidden])
            lab = lambda df_, i: f"{df_.loc[i, 'nome']} {df_.loc[i, 'cognome']}".strip()
            rows = []
            for i in new.index:
                if i not in old.index:
                    rows.append([now, who, i, c_sel, "(nuovo contatto) " + lab(new, i), "", ""])
                else:
                    rows += [[now, who, i, c_sel, f"contatto {lab(new, i)}: {c}", norm(old.loc[i, c]), norm(new.loc[i, c])]
                             for c in CON_COLS[2:] if not same(old.loc[i, c], new.loc[i, c])]
            rows += [[now, who, i, c_sel, "(contatto eliminato) " + lab(old, i), "", ""]
                     for i in old.index.difference(new.index)]
            if rows:
                out = pd.concat([fresh[fresh["cliente"] != c_sel], new.reset_index()[CON_COLS]], ignore_index=True)
                write("Contatti", out, CON_COLS)
                log(rows)
                st.session_state["conv"] = st.session_state.get("conv", 0) + 1
                st.session_state["flash_ana"] = f"Contatti di {c_sel} salvati: {len(rows)} modifiche registrate."
                st.rerun()
            else:
                st.info("Nessuna modifica da salvare.")


# ---------- Anagrafiche ----------
if view == "Anagrafiche":
    if "flash_ana" in st.session_state:
        st.success(st.session_state.pop("flash_ana"))
    # ---------- Consulta: clicca su una riga e vedi tutti gli ordini collegati ----------
    st.markdown("### 🔎 Consulta")
    st.caption("Clicca su una riga di Brand, Stagioni o Clienti per vedere tutti gli ordini collegati.")
    o_all = orders.copy()
    o_all["Pagamento"] = o_all["pay_status"].map(PAY_LBL)
    o_all["Spedizione"] = o_all["ship_status"].map(SHIP_LBL)
    ORD_COLS = ["brand", "stagione", "cliente", "trade_name", "city", "qty", "pct", "valuta",
                "Pagamento", "pay_due", "Spedizione"]
    ORD_CFG = {"brand": "Brand", "stagione": "Stagione", "cliente": "Cliente", "trade_name": "Trade name",
               "city": "Città", "qty": "Qtà", "valuta": "Val.",
               "pct": st.column_config.NumberColumn("Importo", format="%.2f"),
               "pay_due": st.column_config.DateColumn("Scad. pag.", format="DD/MM/YYYY")}

    def show_orders(sub, title):
        st.markdown(f"**{title}**: {len(sub)} ordini · Importo {cur_text(sub, 'pct')} · "
                    f"Qtà {int(sub['qty'].fillna(0).sum())}")
        st.dataframe(sub.sort_values(["brand", "stagione", "cliente", "trade_name"])[ORD_COLS],
                     hide_index=True, use_container_width=True, column_config=ORD_CFG)

    cons = st.radio("Consulta per", ["Brand", "Stagioni", "Clienti"], horizontal=True,
                key="cons_tab", label_visibility="collapsed")

    if cons == "Brand":
        bl = pd.DataFrame({"brand": sort_names(list(brands) + list(orders["brand"]))})
        bl["stagioni"] = bl["brand"].map(lambda b: int(orders.loc[orders["brand"] == b, "stagione"].nunique()))
        bl["ordini"] = bl["brand"].map(lambda b: int((orders["brand"] == b).sum()))
        sb = st.dataframe(bl, hide_index=True, use_container_width=True, on_select="rerun",
                          selection_mode="single-row", key="cons_brand",
                          column_config={"brand": "Brand", "stagioni": "Stagioni", "ordini": "Ordini"})
        if sb.selection.rows:
            b_sel = bl.iloc[sb.selection.rows[0]]["brand"]
            show_orders(o_all[o_all["brand"] == b_sel], b_sel)

    if cons == "Stagioni":
        pairs = {(r.brand, r.code) for r in seasons.itertuples() if str(r.brand).strip() and str(r.code).strip()}
        pairs |= {(r.brand, r.stagione) for r in orders[["brand", "stagione"]].drop_duplicates().itertuples()
                  if str(r.brand).strip() and str(r.stagione).strip()}
        sl = pd.DataFrame(sorted(pairs, key=lambda t: (t[0].casefold(), tuple(-x for x in season_key(t[1])))),
                          columns=["brand", "stagione"])
        sl["ordini"] = [int(((orders["brand"] == b) & (orders["stagione"] == c)).sum())
                        for b, c in zip(sl["brand"], sl["stagione"])]
        ss = st.dataframe(sl, hide_index=True, use_container_width=True, on_select="rerun",
                          selection_mode="single-row", key="cons_season",
                          column_config={"brand": "Brand", "stagione": "Stagione", "ordini": "Ordini"})
        if ss.selection.rows:
            r_ = sl.iloc[ss.selection.rows[0]]
            sub = o_all[(o_all["brand"] == r_["brand"]) & (o_all["stagione"] == r_["stagione"])]
            show_orders(sub, f"{r_['brand']} {r_['stagione']}")
            if st.button("Apri in Ordini", key="cons_open"):
                st.session_state["_open"] = (r_["brand"], r_["stagione"])
                st.rerun()

    if cons == "Clienti":
        cl = pd.DataFrame({"cliente": sort_names(list(clients) + list(orders["cliente"]))})
        cl["ordini"] = cl["cliente"].map(lambda c: int((orders["cliente"] == c).sum()))
        cl["brand"] = cl["cliente"].map(lambda c: ", ".join(sorted(set(orders.loc[orders["cliente"] == c, "brand"]))))
        sc = st.dataframe(cl, hide_index=True, use_container_width=True, on_select="rerun",
                          selection_mode="single-row", key="cons_client",
                          column_config={"cliente": "Cliente", "ordini": "Ordini", "brand": "Brand"})
        if sc.selection.rows:
            c_sel = cl.iloc[sc.selection.rows[0]]["cliente"]
            scheda_cliente(c_sel, who)
            show_orders(o_all[o_all["cliente"] == c_sel], c_sel)

    st.divider()
    st.markdown("### ✏️ Modifica elenchi")
    st.caption("Elenchi usati nei menu. Aggiungi righe in fondo e premi Salva.")
           for name, cols in ANA.items():
        if name != cons:  # mostra solo la sezione scelta nel selettore
            continue
        st.subheader(name)
        data = read(name, tuple(cols))
        if name == "Clienti":  # sempre in ordine alfabetico
            data = pd.DataFrame({"name": sort_names(data["name"])})
        out = st.data_editor(data, num_rows="dynamic", hide_index=True, key=f"ana-{name}")
        if st.button(f"Salva {name.lower()}", key=f"s-{name}"):
            out = out.dropna(how="all")
            if name == "Clienti":
                out = pd.DataFrame({"name": sort_names(out["name"].fillna(""))})
            write(name, out, cols)
            st.rerun()

        if name == "Clienti":
            st.markdown("##### Clienti scritti in modo diverso")
            st.caption("Cerca lo stesso cliente scritto con maiuscole o spazi diversi (es. «FLEUR DE PARIS» e "
                       "«FLEUR de PARIS»). Scegli il nome da tenere: gli ordini verranno aggiornati e il doppione "
                       "sparisce dall'elenco.")
            raw_names = ([str(n) for n in read("Clienti", tuple(cols))["name"] if str(n).strip()]
                         + [str(c) for c in orders["cliente"] if str(c).strip()])
            groups = {}
            for n_ in raw_names:
                groups.setdefault(" ".join(n_.split()).casefold(), set()).add(n_)
            dups = {k_: v_ for k_, v_ in groups.items() if len(v_) > 1}
            n_by = orders["cliente"].map(lambda x: " ".join(str(x).split())).value_counts()
            if not dups:
                st.success("Nessun doppione trovato.")
            for k_, variants in sorted(dups.items()):
                shown = sorted({" ".join(x.split()) for x in variants}, key=lambda x: (x.casefold(), x))
                keep = st.radio("Nome da tenere", shown, key=f"dup_keep_{k_}", horizontal=True,
                                format_func=lambda sp: f"{sp}  ({int(n_by.get(sp, 0))} ordini)")
                if st.button(f"Unisci in «{keep}»", key=f"dup_btn_{k_}"):
                    if not who.strip():
                        st.error("Scrivi il tuo nome nella barra laterale prima di unire.")
                    else:
                        now = datetime.now().strftime("%Y-%m-%d %H:%M")
                        fresh = load_orders(fresh=True)
                        hit = fresh["cliente"].isin(list(variants)) & (fresh["cliente"] != keep)
                        n_upd = int(hit.sum())
                        if n_upd:
                            rows_log = [[now, who, r.id, r.trade_name, "cliente", r.cliente, keep]
                                        for r in fresh[hit].itertuples()]
                            fresh.loc[hit, "cliente"] = keep
                            fresh.loc[hit, ["updated_at", "updated_by"]] = [now, who]
                            save_orders(fresh[COLS])
                            log(rows_log)
                        reg = [n_ for n_ in read("Clienti", tuple(cols))["name"] if str(n_) not in variants] + [keep]
                        write("Clienti", pd.DataFrame({"name": sort_names(reg)}), ANA["Clienti"])
                        st.session_state["flash_ana"] = f"Uniti in «{keep}»: aggiornati {n_upd} ordini."
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
