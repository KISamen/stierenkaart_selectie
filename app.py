import os
os.environ["STREAMLIT_WATCH"] = "false"

import streamlit as st
import pandas as pd
import io

# -------------------------------------------------------
# Mapping table NL
# -------------------------------------------------------
mapping_table_nl = [
    {"Stierenkaart": "superbevruchter", "Titel in bestand": "Superbevruchter", "Formule": None},
    {"Stierenkaart": "ki-code", "Titel in bestand": "Stiercode NL / KI code", "Formule": None},
    {"Stierenkaart": "naam", "Titel in bestand": "Afkorting stier (zoeknaam)", "Formule": None},
    {"Stierenkaart": "vader", "Titel in bestand": "Roepnaam Vader", "Formule": None},
    {"Stierenkaart": "vaders vader", "Titel in bestand": "Roepnaam Moeders Vader", "Formule": None},
    {"Stierenkaart": "PFW", "Titel in bestand": "PFW code", "Formule": None},
    {"Stierenkaart": "aAa", "Titel in bestand": "AAa code", "Formule": None},
    {"Stierenkaart": "Beta caseine", "Titel in bestand": "Betacasine", "Formule": None},
    {"Stierenkaart": "Kappa caseine", "Titel in bestand": "Kappa-caseine", "Formule": None},
    {"Stierenkaart": "prijs", "Titel in bestand": "Prijs", "Formule": None},
    {"Stierenkaart": "prijs gesekst", "Titel in bestand": "", "Formule": None},
    {"Stierenkaart": "&betrouwbaarheid productie", "Titel in bestand": "Official Production Evaluation in this Country %betrouwbaarheid (Productie-index)", "Formule": None},
    {"Stierenkaart": "kg melk", "Titel in bestand": "Official Production Evalution in this Country KG Melk", "Formule": "/10"},
    {"Stierenkaart": "%vet", "Titel in bestand": "Offical Production Evaluation in this Country %vet", "Formule": "/100"},
    {"Stierenkaart": "%eiwit", "Titel in bestand": "Official Production Evaluation in this County %eiwit", "Formule": "/100"},
    {"Stierenkaart": "kg vet", "Titel in bestand": "Official Production Evaluation in this Country KG vet", "Formule": "/10"},
    {"Stierenkaart": "kg eiwit", "Titel in bestand": "Official Production Evaluation in this Country KG eiwit", "Formule": None},
    {"Stierenkaart": "INET", "Titel in bestand": "Official Production Evaluation in this Country Inet", "Formule": None},
    {"Stierenkaart": "NVI", "Titel in bestand": "Official Production Evaluation in this Country NVI", "Formule": None},
    {"Stierenkaart": "TIP", "Titel in bestand": "", "Formule": None},
    {"Stierenkaart": "%betrouwbaarheid exterieur", "Titel in bestand": "%betrouwbaarheid (exterieur-index)", "Formule": None},
    {"Stierenkaart": "frame", "Titel in bestand": "GENERAL CHARACTERISTICS frame", "Formule": "/100"},
    {"Stierenkaart": "uier", "Titel in bestand": "GENERAL CHARACTERISTICS uier", "Formule": "/100"},
    {"Stierenkaart": "benen", "Titel in bestand": "GENERAL CHARACTERISTICS benen", "Formule": "/100"},
    {"Stierenkaart": "totaal", "Titel in bestand": "GENERAL CHARACTERISTICS totaal (Exterieur-index)", "Formule": "/100"},
    {"Stierenkaart": "hoogtemaat", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY hoogtemaat", "Formule": "/100"},
    {"Stierenkaart": "voorhand", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorhand", "Formule": "/100"},
    {"Stierenkaart": "inhoud", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY inhoud", "Formule": "/100"},
    {"Stierenkaart": "ribvorm", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY openheid", "Formule": "/100"},
    {"Stierenkaart": "conditiescore", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY conditiescore", "Formule": "/100"},
    {"Stierenkaart": "kruisligging", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY kruisligging", "Formule": "/100"},
    {"Stierenkaart": "kruisbreedte", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY kruisbreedte", "Formule": "/100"},
    {"Stierenkaart": "beenstand achter", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achteraanzicht benen", "Formule": "/100"},
    {"Stierenkaart": "beenstand zij", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY beenstand zij", "Formule": "/100"},
    {"Stierenkaart": "klauwhoek", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY klauwhoek", "Formule": "/100"},
    {"Stierenkaart": "voorbeenstand", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorbeenstand", "Formule": "/100"},
    {"Stierenkaart": "beengebruik", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY beengebruik", "Formule": "/100"},
    {"Stierenkaart": "vooruieraanhechting", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY vooruieraanhechting", "Formule": "/100"},
    {"Stierenkaart": "voorspeenplaatsing", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorspeenplaatsing", "Formule": "/100"},
    {"Stierenkaart": "speenlengte", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY speenlengte", "Formule": "/100"},
    {"Stierenkaart": "uierdiepte", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY uierdiepte", "Formule": "/100"},
    {"Stierenkaart": "achteruierhoogte", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achteruierhoogte", "Formule": "/100"},
    {"Stierenkaart": "ophangband", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY ophangband", "Formule": "/100"},
    {"Stierenkaart": "achterspeenplaatsing", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achterspeenplaatsing", "Formule": "/100"},
    {"Stierenkaart": "geboortegemak", "Titel in bestand": "OFFICIAL CALVING EASE EVALUATION IN THIS COUNTRY geboortegemak", "Formule": "/100"},
    {"Stierenkaart": "melksnelheid", "Titel in bestand": "OFFICIAL MILKING SPEED AND TEMPERAMENT EVALUATION IN THIS COUNTRY melksnelheid", "Formule": "/100"},
    {"Stierenkaart": "celgetal", "Titel in bestand": "OFFICIAL SOMATIC CELL COUNT EVALUATION IN THIS COUNTRY celgetal", "Formule": "/100"},
    {"Stierenkaart": "vruchtbaarheid", "Titel in bestand": "OFFICIAL FEMALE FERTILITY EVALUATION IN THIS COUNTRY vruchtbaarheid", "Formule": "/100"},
    {"Stierenkaart": "karakter", "Titel in bestand": "OFFICIAL MILKING SPEED AND TEMPERAMENT EVALUATION IN THIS COUNTRY karakter", "Formule": "/100"},
    {"Stierenkaart": "laatrijpheid", "Titel in bestand": "OFFICIAL CALVING EASE EVALUATION IN THIS COUNTRY laatrijpheid", "Formule": "/100"},
    {"Stierenkaart": "persistentie", "Titel in bestand": "", "Formule": "/100"},
    {"Stierenkaart": "klauwgezondheid", "Titel in bestand": "OFFICIAL CLAW HEALTH EVALUATION IN THIS COUNTRY klauwgezondheid", "Formule": "/100"},
    {"Stierenkaart": "levensduur", "Titel in bestand": "OFFICIAL CALF LIVABILITY EVALUATION IN THIS COUNTRY levensduur", "Formule": "/100"}
]

# -------------------------------------------------------
# Mapping table Canada
# -------------------------------------------------------
mapping_table_canada = [
    {"Stierenkaart": "ki-code", "Titel in bestand": "Stiercode NL / KI code", "Formule": None},
    {"Stierenkaart": "Name", "Titel in bestand": "Afkorting stier (zoeknaam)", "Formule": None},
    {"Stierenkaart": "Pedigree sire", "Titel in bestand": "Roepnaam Vader", "Formule": None},
    {"Stierenkaart": "Pedigree mat grandsire", "Titel in bestand": "Roepnaam Vaders Vader", "Formule": None},
    {"Stierenkaart": "aAa", "Titel in bestand": "AAa code", "Formule": None},
    {"Stierenkaart": "prijs", "Titel in bestand": "Prijs", "Formule": None},
    {"Stierenkaart": "prijs gesekst", "Titel in bestand": "", "Formule": None},
    {"Stierenkaart": "%reliability", "Titel in bestand": "Official Production Evaluation in this Country %betrouwbaarheid (Productie-index)", "Formule": None},
    {"Stierenkaart": "kg milk", "Titel in bestand": "Official Production Evalution in this Country KG Melk", "Formule": "/10"},
    {"Stierenkaart": "%fat", "Titel in bestand": "Offical Production Evaluation in this Country %vet", "Formule": "/100"},
    {"Stierenkaart": "%protein", "Titel in bestand": "Official Production Evaluation in this County %eiwit", "Formule": "/100"},
    {"Stierenkaart": "kg fat", "Titel in bestand": "Official Production Evaluation in this Country KG vet", "Formule": "/10"},
    {"Stierenkaart": "kg protein", "Titel in bestand": "Official Production Evaluation in this Country KG eiwit", "Formule": None},
    {"Stierenkaart": "%reliability conformation traits", "Titel in bestand": "%betrouwbaarheid (exterieur-index)", "Formule": None},
    {"Stierenkaart": "frame", "Titel in bestand": "GENERAL CHARACTERISTICS frame", "Formule": "/100"},
    {"Stierenkaart": "udder", "Titel in bestand": "GENERAL CHARACTERISTICS uier", "Formule": "/100"},
    {"Stierenkaart": "feet & legs", "Titel in bestand": "GENERAL CHARACTERISTICS benen", "Formule": "/100"},
    {"Stierenkaart": "final score", "Titel in bestand": "GENERAL CHARACTERISTICS totaal (Exterieur-index)", "Formule": "/100"},
    {"Stierenkaart": "stature", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY hoogtemaat", "Formule": "/100"},
    {"Stierenkaart": "chestwidth", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorhand", "Formule": "/100"},
    {"Stierenkaart": "body depth", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY inhoud", "Formule": "/100"},
    {"Stierenkaart": "angularity", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY openheid", "Formule": "/100"},
    {"Stierenkaart": "condition score", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY conditiescore", "Formule": "/100"},
    {"Stierenkaart": "rump angle", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY kruisligging", "Formule": "/100"},
    {"Stierenkaart": "rump width", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY kruisbreedte", "Formule": "/100"},
    {"Stierenkaart": "rear legs rear view", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achteraanzicht benen", "Formule": "/100"},
    {"Stierenkaart": "rear leg set", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY beenstand zij", "Formule": "/100"},
    {"Stierenkaart": "foot angle", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY klauwhoek", "Formule": "/100"},
    {"Stierenkaart": "front feet orientation", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorbeenstand", "Formule": "/100"},
    {"Stierenkaart": "locomotion", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY beengebruik", "Formule": "/100"},
    {"Stierenkaart": "fore udder attachment", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY vooruieraanhechting", "Formule": "/100"},
    {"Stierenkaart": "fore teat placement", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorspeenplaatsing", "Formule": "/100"},
    {"Stierenkaart": "teat length", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY speenlengte", "Formule": "/100"},
    {"Stierenkaart": "udder depth", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY uierdiepte", "Formule": "/100"},
    {"Stierenkaart": "rear udder height", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achteruierhoogte", "Formule": "/100"},
    {"Stierenkaart": "central ligament", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY ophangband", "Formule": "/100"},
    {"Stierenkaart": "rear teat placement", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achterspeenplaatsing", "Formule": "/100"},
    {"Stierenkaart": "persistency", "Titel in bestand": "", "Formule": "/100"},
    {"Stierenkaart": "calving ease", "Titel in bestand": "OFFICIAL CALVING EASE EVALUATION IN THIS COUNTRY geboortegemak", "Formule": "/100"},
    {"Stierenkaart": "milking speed", "Titel in bestand": "OFFICIAL MILKING SPEED AND TEMPERAMENT EVALUATION IN THIS COUNTRY melksnelheid", "Formule": "/100"},
    {"Stierenkaart": "somatic cell count", "Titel in bestand": "OFFICIAL SOMATIC CELL COUNT EVALUATION IN THIS COUNTRY celgetal", "Formule": "/100"},
    {"Stierenkaart": "female fertility", "Titel in bestand": "OFFICIAL FEMALE FERTILITY EVALUATION IN THIS COUNTRY vruchtbaarheid", "Formule": "/100"},
    {"Stierenkaart": "temperament", "Titel in bestand": "OFFICIAL MILKING SPEED AND TEMPERAMENT EVALUATION IN THIS COUNTRY karakter", "Formule": "/100"},
    {"Stierenkaart": "maturity rate", "Titel in bestand": "OFFICIAL CALVING EASE EVALUATION IN THIS COUNTRY laatrijpheid", "Formule": "/100"},
    {"Stierenkaart": "hoofhealth", "Titel in bestand": "OFFICIAL CLAW HEALTH EVALUATION IN THIS COUNTRY klauwgezondheid", "Formule": "/100"},
    {"Stierenkaart": "Beta caseine", "Titel in bestand": "Betacasine", "Formule": None},
    {"Stierenkaart": "Kappa caseine", "Titel in bestand": "Kappa-caseine", "Formule": None},
    {"Stierenkaart": "Superbevruchter", "Titel in bestand": "Superbevruchter", "Formule": None}
]
# -------------------------------------------------------
# Functie: Excel laden
# -------------------------------------------------------
def load_excel(file):
    try:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Fout bij laden Excel: {e}")
        return None

# -------------------------------------------------------
# Functie: sorteer stieren op ras en naam
# -------------------------------------------------------
def custom_sort_ras(df):
    if "Ras" not in df.columns:
        df["Ras"] = ""
    naam_col = "naam" if "naam" in df.columns else ("Name" if "Name" in df.columns else None)
    if naam_col and naam_col not in df.columns:
        df[naam_col] = ""
    order_map = {"Holstein zwartbont": 1, "Red Holstein": 2}
    df["ras_sort"] = df["Ras"].map(order_map).fillna(3)
    df_sorted = df.sort_values(by=["ras_sort", naam_col], ascending=True)
    df_sorted.drop(columns=["ras_sort"], inplace=True)
    return df_sorted

# -------------------------------------------------------
# Functie: maak Top 5-tabellen (alleen NL)
# -------------------------------------------------------
def create_top5_table(df):
    fokwaarden = [
        "geboortegemak",
        "celgetal",
        "vruchtbaarheid",
        "klauwgezondheid",
        "uier",
        "benen"
    ]
    blocks = []
    if df.empty or "Ras" not in df.columns:
        return pd.DataFrame()
    # filter alleen NL-rassen
    df_nl = df[df["Ras"].isin(["Holstein zwartbont", "Holstein zwartbont + RF", "Red Holstein"])].copy()
    for fok in fokwaarden:
        if fok not in df_nl.columns:
            df_nl[fok] = pd.NA
        header = {
            "Fokwaarde": fok,
            "zwartbont_stier": "Stier",
            "zwartbont_value": "Waarde",
            "roodbont_stier": "Stier",
            "roodbont_value": "Waarde"
        }
        blocks.append(header)
        df_z = df_nl[df_nl["Ras"] == "Holstein zwartbont"].copy()
        df_z[fok] = pd.to_numeric(df_z[fok], errors='coerce')
        df_z = df_z.sort_values(by=fok, ascending=False)
        df_r = df_nl[df_nl["Ras"] == "Red Holstein"].copy()
        df_r[fok] = pd.to_numeric(df_r[fok], errors='coerce')
        df_r = df_r.sort_values(by=fok, ascending=False)
        for i in range(5):
            row = {k: "" for k in header}
            if i < len(df_z):
                row["zwartbont_stier"] = str(df_z.iloc[i]["naam"])
                row["zwartbont_value"] = str(df_z.iloc[i][fok])
            if i < len(df_r):
                row["roodbont_stier"] = str(df_r.iloc[i]["naam"])
                row["roodbont_value"] = str(df_r.iloc[i][fok])
            blocks.append(row)
        # lege rij na blok
        blocks.append({k: "" for k in header})
    return pd.DataFrame(blocks)

# -------------------------------------------------------
# Hoofdapplicatie
# -------------------------------------------------------
def main():
    st.set_page_config(layout="wide")
    st.title("Stierenkaart Generator")

    # Taalkeuze
    taal = st.selectbox("Kies stierenkaart type:", ["Nederland", "Canada"])
    mapping_table = mapping_table_nl if taal == "Nederland" else mapping_table_canada

    uploaded_file = st.file_uploader("Upload Excel bestand", type=["xlsx"])
    if not uploaded_file:
        return

    df_raw = load_excel(uploaded_file)
    if df_raw is None:
        return

    # Bouw final_data
    final_data = {}
    for mapping in mapping_table:
        titel = mapping["Titel in bestand"]
        std_naam = mapping["Stierenkaart"]
        formule = mapping["Formule"]
        if titel and titel in df_raw.columns:
            kolom = df_raw[titel].replace([99999, "+999"], pd.NA)
            if formule:
                kolom = pd.to_numeric(kolom, errors="coerce")
                if formule == "/10":
                    kolom = kolom / 10
                elif formule == "/100":
                    kolom = kolom / 100
            final_data[std_naam] = kolom
        else:
            final_data[std_naam] = ""

    df_mapped = pd.DataFrame(final_data)

    # Pinkenstier alleen voor NL
    if taal == "Nederland" and "geboortegemak" in df_mapped.columns:
        df_mapped["pinkenstier"] = df_mapped["geboortegemak"].apply(
            lambda x: "p" if pd.notna(x) and x > 1.0 else ""
        )
    else:
        df_mapped["pinkenstier"] = ""

    # Dynamische sleutelkolommen
    naam_col = "naam" if taal == "Nederland" else "Name"
    ki_col = "ki-code"

    # Kolomvolgorde
    volgorde = ["superbevruchter", ki_col, naam_col, "pinkenstier"] + [m["Stierenkaart"] for m in mapping_table if m["Stierenkaart"] not in ["superbevruchter", ki_col, naam_col, "pinkenstier"]]
    volgorde = [c for c in volgorde if c in df_mapped.columns]
    overige = [c for c in df_mapped.columns if c not in volgorde]
    df_mapped = df_mapped[volgorde + overige]

    # Display-kolom
    if ki_col in df_mapped.columns and naam_col in df_mapped.columns:
        df_mapped["Display"] = df_mapped[ki_col].astype(str) + " - " + df_mapped[naam_col].astype(str)

        keuze = st.multiselect("Selecteer stieren:", df_mapped["Display"].tolist(), default=None)
        if keuze:
            codes = [x.split(" - ")[0] for x in keuze]
            df_selected = df_mapped[df_mapped[ki_col].isin(codes)].copy()
            df_selected = custom_sort_ras(df_selected)

            st.subheader("Stierenkaart")
            st.dataframe(df_selected, use_container_width=True)

            # Top-5 alleen NL
            if taal == "Nederland":
                df_top5 = create_top5_table(df_selected)
                st.subheader("Top 5-tabellen")
                st.dataframe(df_top5, use_container_width=True)

            # Download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_selected.to_excel(writer, sheet_name='Stierenkaart', index=False)
                if taal == "Nederland":
                    df_top5.to_excel(writer, sheet_name='Top5', index=False)
            st.download_button("Download Excel", data=output.getvalue(), file_name="stierenkaart.xlsx")
        else:
            st.info("Selecteer één of meer stieren om de gegevens te zien en te downloaden.")
    else:
        st.warning(f"Kolommen '{ki_col}' en/of '{naam_col}' ontbreken in de gemapte data.")

if __name__ == "__main__":
    main()
