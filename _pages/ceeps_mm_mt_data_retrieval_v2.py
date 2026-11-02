"""CEEPS: oddaja posamičnih zahtev za več merilnih točk.

Zagon: streamlit run ceeps_mm_mt_data_retrieval_v2.py
Nastavitev: .streamlit/secrets.toml -> encoded_string_nme = "..."
"""

import base64
import hashlib
import io
import json
import re
import time
from datetime import datetime

import pandas as pd
import requests
import streamlit as st

API_URL = (
    "https://api.informatika.si/enotna-vstopna-tocka/"
    "evidenca-zahtev/merilne-tocke/podatki-mt-mm-meritve"
)
VRSTE_DOKUMENTOV = [
    "POOBLASTILO",
    "POOBLASTILO_ZA_PRIDOBITEV_MERILNIH_PODATKOV",
    "VLOGA_ZA_PRIKLJUCITEV_IN_DOSTOP_DO_OMREZJA",
    "DRUGA",
]
GSRN_RE = re.compile(r"^[0-9]{18}$")
MAX_TOCK = 500
MAX_PDF_BYTES = 10 * 1024 * 1024  # Lokalna zaščitna omejitev; ni navedba omejitve CEEPS.


def razdeli_rocni_vnos(vnos: str) -> list[str]:
    return [v.strip() for v in re.split(r"[,;\n\r\t]+", vnos) if v.strip()]


def pridobi_excel(datoteka) -> list[str]:
    """Prvi stolpec iz prvega lista; prepozna poljubno naslovno vrstico.

    Pozor: Excel številk z 18 števkami lahko nepopravljivo zaokroži.
    Zato je treba v Excelu stolpec pred vnosom nastaviti na 'Besedilo'.
    """
    podatki = pd.read_excel(
        io.BytesIO(datoteka.getvalue()),
        header=None,
        sheet_name=0,
        dtype=str,
        keep_default_na=False,
        engine="openpyxl",
    )
    if podatki.empty or podatki.shape[1] == 0:
        return []
    izhod = []
    for vrednost in podatki.iloc[:, 0].tolist():
        v = str(vrednost).strip()
        if not v:
            continue
        if v.lower() in {"gsrnmt", "gsrn mt", "merilna točka", "merilna tocka", "gsrn"}:
            continue
        izhod.append(v)
    return izhod


def preveri_gsrn(seznam: list[str]):
    pravilni, napacni, podvojeni = [], [], []
    ze_videni = set()
    for vrednost in seznam:
        if not GSRN_RE.fullmatch(vrednost):
            napacni.append(vrednost)
        elif vrednost in ze_videni:
            podvojeni.append(vrednost)
        else:
            pravilni.append(vrednost)
            ze_videni.add(vrednost)
    return pravilni, napacni, podvojeni


def pripravi_telo(gsrn, placnik, ima_pooblastilo, priloga):
    return {
        "gsrnMT": gsrn,
        "placnikAliProdajalec": placnik,
        "imaVeljavnoPooblastiloZaPridobivanjePodatkov": ima_pooblastilo,
        "priloge": [priloga],
    }


def rezultat_vrstica(gsrn, status, http_koda=None, id_zahteva=None, opomba=""):
    return {
        "Merilna točka (GSRN MT)": gsrn,
        "Status": status,
        "HTTP koda": http_koda,
        "ID zahteve": id_zahteva,
        "Podrobnosti": opomba,
        "Čas obdelave": datetime.now().astimezone().isoformat(timespec="seconds"),
    }


def poslji_eno(seja, headers, gsrn, placnik, ima_pooblastilo, priloga):
    telo = pripravi_telo(gsrn, placnik, ima_pooblastilo, priloga)
    try:
        odziv = seja.post(API_URL, headers=headers, json=telo, timeout=(15, 120))
        if odziv.status_code == 200:
            try:
                odgovor = odziv.json()
                id_zahteva = odgovor.get("idZahteva") if isinstance(odgovor, dict) else None
            except ValueError:
                id_zahteva = None
            return rezultat_vrstica(gsrn, "Uspešno", 200, id_zahteva)
        if odziv.status_code == 401:
            return rezultat_vrstica(gsrn, "Napaka", 401, opomba="Neavtoriziran dostop")
        if odziv.status_code == 403:
            return rezultat_vrstica(gsrn, "Napaka", 403, opomba="Dostop zavrnjen")
        # Ne izpisujemo response.text: lahko vsebuje občutljive osebne podatke.
        return rezultat_vrstica(
            gsrn, "Napaka", odziv.status_code,
            opomba="CEEPS je vrnil neuspešen HTTP status. Preverite zahtevo."
        )
    except requests.exceptions.Timeout:
        # Po timeoutu ni mogoče ugotoviti, ali je strežnik zahtevo že sprejel.
        return rezultat_vrstica(gsrn, "Neznano", opomba="Časovna omejitev; pred ponovitvijo preverite v CEEPS")
    except requests.exceptions.RequestException:
        return rezultat_vrstica(gsrn, "Neznano", opomba="Omrežna napaka; pred ponovitvijo preverite v CEEPS")


def prstni_odtis(gsrni, placnik, pooblastilo, priloga):
    material = json.dumps(
        {"gsrni": gsrni, "placnik": placnik, "pooblastilo": pooblastilo,
         "naziv_priloge": priloga["naziv"], "tip_priloge": priloga["vrstaDokumenta"],
         "hash_priloge": hashlib.sha256(priloga["datoteka"].encode()).hexdigest()},
        ensure_ascii=False, sort_keys=True,
    )
    return hashlib.sha256(material.encode()).hexdigest()


def main():
    st.set_page_config(page_title="CEEPS | Zahteve za MT/MM meritve", page_icon="⚡", layout="wide")
    st.title("Zahteva za podatke MT/MM meritev")
    st.caption("CEEPS · Posamične zahteve za eno ali več merilnih točk z enakimi podatki")

    st.subheader("1. Podatki o merilnih točkah")
    nacin = st.radio("Način vnosa", ["Ročni vnos", "Excel (.xlsx)"], horizontal=True)
    tocke = []
    if nacin == "Ročni vnos":
        besedilo = st.text_area(
            "GSRN MT – ena merilna točka v vrstici",
            placeholder="483111589999999999\n483111589999999998",
            height=145,
            help="Dovoljeni so tudi ločilniki vejica, podpičje ali tabulator.",
        )
        tocke = razdeli_rocni_vnos(besedilo)
    else:
        st.info("V Excelu nastavite prvi stolpec na obliko **Besedilo** še pred vnosom 18-mestnih številk. Excel sicer lahko spremeni zadnje števke.")
        xlsx = st.file_uploader("Naložite Excel (prvi stolpec, prvi list)", type=["xlsx"])
        if xlsx is not None:
            try:
                tocke = pridobi_excel(xlsx)
            except Exception:
                st.error("Datoteke ni mogoče prebrati. Preverite, da gre za veljaven .xlsx in da so GSRN-ji v prvem stolpcu.")

    pravilni, napacni, podvojeni = preveri_gsrn(tocke)
    a, b, c = st.columns(3)
    a.metric("Veljavnih MT", len(pravilni))
    b.metric("Neveljavnih vnosov", len(napacni))
    c.metric("Podvojenih vnosov", len(podvojeni))
    if napacni:
        st.error("Neveljavni GSRN MT (zahtevano natanko 18 števk): " + ", ".join(napacni[:12]))
    if podvojeni:
        st.warning("Podvojeni GSRN MT bodo izločeni: " + ", ".join(dict.fromkeys(podvojeni[:12])))
    if pravilni:
        with st.expander(f"Pregled merilnih točk ({len(pravilni)})", expanded=False):
            st.dataframe(pd.DataFrame({"GSRN MT": pravilni}), use_container_width=True, hide_index=True)
    if len(pravilni) > MAX_TOCK:
        st.error(f"Naenkrat je dovoljenih največ {MAX_TOCK} točk. Seznam razdelite v več sklopov.")

    st.subheader("2. Podatki o plačniku / prodajalcu")
    col1, col2 = st.columns(2)
    with col1:
        naziv = st.text_input("Naziv", value="")
        ulica = st.text_input("Ulica", value="")
        hisna_stevilka = st.text_input("Hišna številka", value="")
        davcni_zavezanec = st.checkbox("Davčni zavezanec", value=False)
    with col2:
        postna_stevilka = st.text_input("Poštna številka", value="")
        posta = st.text_input("Pošta", value="")
        davcna_stevilka = st.text_input("Davčna številka", value="")
    placnik = {
        "naziv": naziv.strip(), "ulica": ulica.strip(),
        "hisnaStevilka": hisna_stevilka.strip(),
        "postnaStevilka": postna_stevilka.strip(), "posta": posta.strip(),
        "davcniZavezanec": davcni_zavezanec, "davcnaStevilka": davcna_stevilka.strip(),
    }

    st.subheader("3. Pooblastilo")
    ima_pooblastilo = st.checkbox("Ima veljavno pooblastilo za pridobivanje podatkov?", value=True)
    st.subheader("4. Priloga (PDF)")
    pdf = st.file_uploader("Naložite PDF pooblastila / dokumenta", type=["pdf"])
    vrsta_dokumenta = st.selectbox("Vrsta dokumenta", VRSTE_DOKUMENTOV)
    pdf_bytes = pdf.getvalue() if pdf is not None else None
    priloga = None
    if pdf_bytes is not None:
        if not pdf_bytes.startswith(b"%PDF-"):
            st.error("Naložena datoteka ni veljaven PDF.")
        elif len(pdf_bytes) > MAX_PDF_BYTES:
            st.error("PDF presega lokalno omejitev 10 MB. Omejitev je varovalka aplikacije.")
        else:
            priloga = {
                "naziv": pdf.name,
                "datoteka": base64.b64encode(pdf_bytes).decode("ascii"),
                "vrstaDokumenta": vrsta_dokumenta,
            }

    st.subheader("5. Pregled in pošiljanje")
    st.info("Za vsak GSRN MT se pošlje **samostojna POST zahteva**. Podatki plačnika, vrsta dokumenta in PDF so pri vseh enaki. Podvojenih vnosov ne pošiljamo.")
    st.caption("Pošiljanje lahko ustavite med posameznimi zahtevami z ustavitvijo izvajanja Streamlit. Že poslane zahteve ostanejo poslane.")
    potrditev = st.checkbox(f"Potrjujem oddajo {len(pravilni)} posameznih zahtev v CEEPS.")
    obvezni = ["naziv", "ulica", "hisnaStevilka", "postnaStevilka", "posta"]
    placnik_veljaven = all(placnik[k] for k in obvezni)
    if davcni_zavezanec and not davcna_stevilka.strip():
        placnik_veljaven = False
    if not placnik_veljaven:
        st.caption("Za pošiljanje izpolnite naziv, ulico, hišno številko, poštno številko in pošto; za davčnega zavezanca tudi davčno številko.")

    dovoljeno = bool(pravilni and not napacni and len(pravilni) <= MAX_TOCK and priloga and placnik_veljaven and potrditev)
    if "zgodovina" not in st.session_state:
        st.session_state.zgodovina = {}
    if "zadnja_serija" not in st.session_state:
        st.session_state.zadnja_serija = None

    if st.button("📤 Pošlji posamezne zahteve", type="primary", disabled=not dovoljeno):
        try:
            kodiran = st.secrets["encoded_string_nme"]
        except (KeyError, FileNotFoundError):
            st.error("V Streamlit secrets manjka ključ encoded_string_nme.")
            return
        if not isinstance(kodiran, str) or not kodiran.strip():
            st.error("Nastavitev encoded_string_nme je prazna.")
            return

        podpis = prstni_odtis(pravilni, placnik, ima_pooblastilo, priloga)
        # Rezultate serije obdržimo v seji, da ponoven klik ne ustvari dvojnikov.
        if podpis in st.session_state.zgodovina:
            st.warning("Ta nabor zahtev je bil v tej seji že obdelan. Rezultati so prikazani spodaj. Za preverjanje statusa uporabite CEEPS.")
            st.session_state.zadnja_serija = podpis
        else:
            headers = {
                "accept": "application/json",
                "Content-Type": "application/json",
                "Authorization": f"Basic {kodiran.strip()}",
            }
            st.session_state.zgodovina[podpis] = []
            st.session_state.zadnja_serija = podpis
            bar = st.progress(0, text="Priprava zahtev ...")
            status = st.empty()
            with requests.Session() as seja:
                for indeks, gsrn in enumerate(pravilni, 1):
                    status.write(f"Pošiljanje {indeks}/{len(pravilni)} · GSRN MT {gsrn}")
                    vrstica = poslji_eno(seja, headers, gsrn, placnik, ima_pooblastilo, priloga)
                    st.session_state.zgodovina[podpis].append(vrstica)
                    bar.progress(indeks / len(pravilni), text=f"Obdelanih {indeks}/{len(pravilni)}")
                    if vrstica["HTTP koda"] in (401, 403):
                        st.warning("Pošiljanje je prekinjeno zaradi težave z dostopom. Preverite pooblastila in prijavo.")
                        break
            status.empty()

    zadnja = st.session_state.zadnja_serija
    if zadnja is not None and zadnja in st.session_state.zgodovina:
        rezultati = st.session_state.zgodovina[zadnja]
        if rezultati:
            st.markdown("### Rezultati zadnjega pošiljanja")
            frame = pd.DataFrame(rezultati)
            uspesni = sum(v["Status"] == "Uspešno" for v in rezultati)
            napake = sum(v["Status"] == "Napaka" for v in rezultati)
            neznano = sum(v["Status"] == "Neznano" for v in rezultati)
            x, y, z = st.columns(3)
            x.metric("Uspešno", uspesni)
            y.metric("Napake", napake)
            z.metric("Neznan rezultat", neznano)
            if neznano:
                st.warning("Zahtev z neznanim rezultatom ne pošiljajte ponovno brez preverjanja, ali jih je CEEPS že prejel.")
            st.dataframe(frame, hide_index=True, use_container_width=True)
            st.download_button(
                "⬇️ Prenesi poročilo CSV", frame.to_csv(index=False, sep=";").encode("utf-8-sig"),
                file_name="ceeps_rezultati_zahtev.csv", mime="text/csv",
            )


if __name__ == "__main__":
    main()
