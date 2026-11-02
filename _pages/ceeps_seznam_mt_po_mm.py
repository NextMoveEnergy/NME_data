"""Streamlit stran: CEEPS GET /merilna-mesta/seznam-mt za več MM.

Vključitev v večstransko aplikacijo: kopirajte v pages/02_Seznam_MT_po_MM.py.
Zahteva: pip install streamlit pandas openpyxl requests
Avtentikacija: .streamlit/secrets.toml -> encoded_string_nme = "..."

POMEMBNO: Swagger na sliki kaže pot /merilna-mesta/seznam-mt, ne pa
celotnega URL strežnika; privzeta osnova je prevzeta iz stare aplikacije.
Preverite jo glede na dejanski Swagger Servers / vaš CEEPS deployment.
"""

import io
import re
import time
from datetime import datetime
from urllib.parse import urlsplit

import pandas as pd
import requests
import streamlit as st

# Osnova iz obstoječe aplikacije; preverite v Swagger Servers.
DEFAULT_API_URL = "https://api.informatika.si/enotna-vstopna-tocka/merilna-mesta/seznam-mt"
MAX_MM = 500
INPUT_COL = "Merilno mesto"


def split_values(text):
    return [v.strip() for v in re.split(r"[\r\n;,\t]+", text or "") if v.strip()]


def read_excel(uploaded):
    # Preberemo prvi stolpec prvega lista kot besedilo; prva vrstica je lahko naslov.
    df = pd.read_excel(io.BytesIO(uploaded.getvalue()), sheet_name=0, header=None,
                       dtype=str, keep_default_na=False, engine="openpyxl")
    if df.empty:
        return []
    values = []
    ignored_headers = {"mm", "gsrnmm", "gsrn mm", "merilno mesto", "merilna mesta",
                       "enotniidentifikatormm", "enotni identifikator mm", "identifikator mm"}
    for v in df.iloc[:, 0]:
        s = str(v).strip()
        if s and s.lower() not in ignored_headers:
            values.append(s)
    return values


def clean_items(items, tip):
    unique, invalid, repeated = [], [], []
    seen = set()
    for x in items:
        # GSRN MM prikaz v Swaggerju je 18-mestni; enotni identifikator je string (npr. 5-4865).
        correct = bool(re.fullmatch(r"[0-9]{18}", x)) if tip == "GSRN MM" else bool(re.fullmatch(r"[A-Za-z0-9_-]+", x))
        if not correct:
            invalid.append(x)
        elif x in seen:
            repeated.append(x)
        else:
            seen.add(x)
            unique.append(x)
    return unique, invalid, repeated


def fetch_one(session, url, secret, mm, tip):
    field = "gsrnMM" if tip == "GSRN MM" else "enotniIdentifikatorMM"
    headers = {"accept": "application/json", "Authorization": f"Basic {secret}"}
    row_base = {INPUT_COL: mm, "Vrsta identifikatorja": tip}
    try:
        r = session.get(url, headers=headers, params={field: mm}, timeout=(15, 45))
        if r.status_code != 200:
            reason = {401: "Neavtoriziran dostop", 403: "Dostop zavrnjen", 404: "Endpoint ni najden"}.get(
                r.status_code, "Neuspešen HTTP odgovor")
            return {**row_base, "Status": "Napaka", "HTTP": r.status_code, "Število MT": 0,
                    "Opomba": reason}, []
        try:
            payload = r.json()
        except ValueError:
            return {**row_base, "Status": "Napaka", "HTTP": 200, "Število MT": 0,
                    "Opomba": "Odgovor ni veljaven JSON"}, []
        # Swagger response: array [{gsrnMT: string, vrstaMT: string}, ...]
        if not isinstance(payload, list):
            return {**row_base, "Status": "Napaka", "HTTP": 200, "Število MT": 0,
                    "Opomba": "Nepričakovana oblika odgovora; pričakovan seznam"}, []
        found = []
        for item in payload:
            if not isinstance(item, dict):
                continue
            gsrn = str(item.get("gsrnMT") or "").strip()
            vrsta = str(item.get("vrstaMT") or "").strip()
            if gsrn:
                found.append({**row_base, "GSRN MT": gsrn, "Vrsta MT (CEEPS)": vrsta})
        return {**row_base, "Status": "Uspešno" if found else "Brez MT", "HTTP": 200,
                "Število MT": len(found), "Opomba": "" if found else "CEEPS je vrnil prazen seznam"}, found
    except requests.exceptions.Timeout:
        return {**row_base, "Status": "Napaka", "HTTP": None, "Število MT": 0,
                "Opomba": "Časovna omejitev povezave"}, []
    except requests.exceptions.RequestException:
        return {**row_base, "Status": "Napaka", "HTTP": None, "Število MT": 0,
                "Opomba": "Napaka omrežne povezave"}, []


def export_excel(status_rows, mt_rows, type_labels=None):
    output = io.BytesIO()
    status = pd.DataFrame(status_rows, columns=[INPUT_COL, "Vrsta identifikatorja", "Status", "HTTP", "Število MT", "Opomba"])
    detail = pd.DataFrame(mt_rows, columns=[INPUT_COL, "Vrsta identifikatorja", "GSRN MT", "Vrsta MT (CEEPS)"])
    type_labels = type_labels or {}
    detail["Pomen vrste MT"] = detail["Vrsta MT (CEEPS)"].map(lambda v: type_labels.get(v, "Nerazvrščeno"))
    # Ena vrstica na MM; ločeni stolpci za odjem/oddajo, več GSRN-jev ločenih z vejico.
    summary = []
    for record in status_rows:
        mm = record[INPUT_COL]
        matched = detail[detail[INPUT_COL] == mm] if not detail.empty else detail
        summary.append({INPUT_COL: mm, "Odjemne MT": ", ".join(matched.loc[matched["Pomen vrste MT"] == "Odjem", "GSRN MT"].tolist()),
                        "Oddajne MT": ", ".join(matched.loc[matched["Pomen vrste MT"] == "Oddaja", "GSRN MT"].tolist()),
                        "Druge / nerazvrščene MT": ", ".join(matched.loc[~matched["Pomen vrste MT"].isin(["Odjem", "Oddaja"]), "GSRN MT"].tolist()),
                        "Status": record["Status"]})
    # Brez podvojenih MT za nadaljnjo oddajo zahtev; ohranimo prvotne GSRN kot str.
    unique = detail[["GSRN MT"]].drop_duplicates().rename(columns={"GSRN MT": "gsrnMT"}) if not detail.empty else pd.DataFrame(columns=["gsrnMT"])
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        status.to_excel(writer, sheet_name="Pregled MM", index=False)
        pd.DataFrame(summary).to_excel(writer, sheet_name="Odjem in oddaja po MM", index=False)
        detail.to_excel(writer, sheet_name="MM in pripadajoce MT", index=False)
        unique.to_excel(writer, sheet_name="MT za oddajo zahtev", index=False)
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for cell in sheet[1]:
                cell.font = __import__("openpyxl").styles.Font(bold=True, color="FFFFFF")
                cell.fill = __import__("openpyxl").styles.PatternFill("solid", fgColor="17365D")
            for col in sheet.columns:
                letter = col[0].column_letter
                maxlen = max((len(str(c.value)) for c in col if c.value is not None), default=12)
                sheet.column_dimensions[letter].width = min(48, max(17, maxlen + 2))
                for cell in list(col)[1:]:
                    if sheet.title in {"MM in pripadajoce MT", "MT za oddajo zahtev", "Pregled MM"}:
                        cell.number_format = "@" if cell.column == 1 or sheet.title == "MT za oddajo zahtev" or (sheet.title == "MM in pripadajoce MT" and cell.column == 3) else cell.number_format
    return output.getvalue()


def main():
    st.set_page_config(page_title="CEEPS | MM → MT", page_icon="⚡", layout="wide")
    st.title("Seznam merilnih točk po merilnih mestih")
    st.caption("CEEPS · GET /merilna-mesta/seznam-mt · Ena poizvedba za vsako merilno mesto")
    st.info("Za vsako MM pridobimo vse pripadajoče MT, skupaj z vrsto `vrstaMT`, kot jo vrne CEEPS. "
            "Ker Swagger na sliki ne podaja pomenov vseh oznak vrste MT, oznak ne preimenujemo na pamet v odjem/oddajo.")

    st.subheader("1. Seznam merilnih mest")
    tip = st.radio("Tip identifikatorja", ["GSRN MM", "Enotni identifikator MM"], horizontal=True,
                   help="Naenkrat uporabljajte en tip identifikatorja; Swagger dovoljuje parametra gsrnMM in enotniIdentifikatorMM.")
    način = st.radio("Način vnosa", ["Ročni vnos", "Excel (.xlsx)"], horizontal=True)
    items = []
    if način == "Ročni vnos":
        sample = "383111589999999999" if tip == "GSRN MM" else "5-4865"
        input_text = st.text_area("Vsako merilno mesto v svojo vrstico", placeholder=f"{sample}\n...", height=145)
        items = split_values(input_text)
    else:
        st.warning("GSRN MM hranite v Excelu kot **besedilo**. Excel lahko 18-mestne številke zaokroži.")
        uploaded = st.file_uploader("Naloži Excel – prvi stolpec, prvi list", type="xlsx")
        if uploaded:
            try:
                items = read_excel(uploaded)
            except Exception:
                st.error("Excela ni mogoče prebrati. Preverite format in prvi stolpec.")

    valid, invalid, duplicates = clean_items(items, tip)
    a, b, c = st.columns(3)
    a.metric("Veljavna MM", len(valid))
    b.metric("Neveljavni vnosi", len(invalid))
    c.metric("Podvojeni vnosi", len(duplicates))
    if invalid:
        st.error("Neveljavni identifikatorji: " + ", ".join(invalid[:10]))
    if duplicates:
        st.warning("Podvojene MM se obravnavajo le enkrat.")
    if valid:
        with st.expander("Predogled seznama MM"):
            st.dataframe(pd.DataFrame({INPUT_COL: valid}), hide_index=True, use_container_width=True)

    st.subheader("2. Povezava in pošiljanje")
    with st.expander("Nastavitve API", expanded=False):
        url = st.text_input("Celoten URL GET storitve", DEFAULT_API_URL,
                            help="Preverite Swagger Servers. API naslov lahko prilagodite, če je drugačen.")
    if not urlsplit(url).scheme == "https" or not urlsplit(url).netloc:
        st.error("URL mora biti veljaven HTTPS naslov.")
    st.caption("Uporabljena je obstoječa Basic avtentikacija iz `st.secrets['encoded_string_nme']`."
               " Vse poizvedbe so GET; ta stran ne oddaja zahtev za merilne podatke.")
    if "ceeps_mm_mt_results" not in st.session_state:
        st.session_state.ceeps_mm_mt_results = None

    disabled = not valid or bool(invalid) or len(valid) > MAX_MM or not url.startswith("https://")
    if len(valid) > MAX_MM:
        st.error(f"Največ {MAX_MM} MM v eni seriji.")
    if st.button(f"Pridobi merilne točke za {len(valid)} MM", type="primary", disabled=disabled):
        secret = st.secrets.get("encoded_string_nme", "")
        if not secret:
            st.error("Manjka nastavitev encoded_string_nme v .streamlit/secrets.toml.")
        else:
            status_rows, mt_rows = [], []
            bar = st.progress(0, text="Pridobivanje podatkov ...")
            status_text = st.empty()
            with requests.Session() as session:
                for index, mm in enumerate(valid):
                    status_text.caption(f"Obdelava {index + 1}/{len(valid)}: {mm}")
                    summary, found = fetch_one(session, url, secret, mm, tip)
                    status_rows.append(summary)
                    mt_rows.extend(found)
                    bar.progress((index + 1) / len(valid))
                    # Ne nadaljujemo z množičnimi klici, če prijava ni veljavna.
                    if summary.get("HTTP") in (401, 403):
                        st.warning("Obdelava je prekinjena zaradi napake avtorizacije.")
                        break
            st.session_state.ceeps_mm_mt_results = {"status": status_rows, "mt": mt_rows,
                                                      "date": datetime.now().strftime("%Y%m%d_%H%M%S")}
            status_text.empty()

    results = st.session_state.ceeps_mm_mt_results
    if results:
        st.subheader("3. Rezultati")
        status_rows, mt_rows = results["status"], results["mt"]
        success_count = sum(row["Status"] == "Uspešno" for row in status_rows)
        no_mt = sum(row["Status"] == "Brez MT" for row in status_rows)
        failed = sum(row["Status"] == "Napaka" for row in status_rows)
        unique_count = len({row["GSRN MT"] for row in mt_rows})
        a, b, c, d = st.columns(4)
        a.metric("MM z MT", success_count)
        b.metric("MM brez MT", no_mt)
        c.metric("Napake", failed)
        d.metric("Unikatne MT", unique_count)
        st.markdown("**Pregled po merilnih mestih**")
        st.dataframe(pd.DataFrame(status_rows), use_container_width=True, hide_index=True)
        st.markdown("**Vse pridobljene merilne točke**")
        if mt_rows:
            st.dataframe(pd.DataFrame(mt_rows), use_container_width=True, hide_index=True)
            vrsta_counts = pd.DataFrame(mt_rows)["Vrsta MT (CEEPS)"].value_counts()
            st.caption("Kode vrst MT iz odgovora: " + ", ".join(f"{k or '(prazno)'}: {v}" for k, v in vrsta_counts.items()))
            st.markdown("**Razvrstitev oznak vrste MT (neobvezno)**")
            st.caption("Swagger ne določa pomena kode OMTO. Potrdite preslikavo s CEEPS dokumentacijo. "
                       "V izvoz se zapišeta ločena stolpca za odjem in oddajo, nerazvrščene točke ostanejo ohranjene.")
            type_labels = {}
            cols = st.columns(min(3, max(1, len(vrsta_counts))))
            for idx, code in enumerate(vrsta_counts.index.tolist()):
                type_labels[code] = cols[idx % len(cols)].selectbox(
                    f"Vrsta MT: {code or '(prazno)'}", ["Nerazvrščeno", "Odjem", "Oddaja", "Drugo"],
                    key=f"ceeps_type_{code}")
        else:
            st.info("Ni pridobljenih merilnih točk.")
        data = export_excel(status_rows, mt_rows, type_labels if mt_rows else {})
        st.download_button("Prenesi rezultate v Excel (.xlsx)", data=data,
                           file_name=f"ceeps_seznam_mt_{results['date']}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        st.caption("Excel vsebuje 4 liste: pregled MM, odjem/oddaja po MM, povezave MM–MT in unikatne GSRN MT, "
                   "ki ga lahko naložite na stran za oddajo zahtev za meritve.")


if __name__ == "__main__":
    main()
