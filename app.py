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

# Mapping table Canada (voorbeeld)
mapping_table_canada = [
    {"Stierenkaart": "ki-code", "Titel in bestand": "Stiercode NL / KI code", "Formule": None},
    {"Stierenkaart": "Name", "Titel in bestand": "Afkorting stier (zoeknaam)", "Formule": None},
    {"Stierenkaart": "Pedigree", "Titel in bestand": "Roepnaam Vader", "Formule": None},
    # voeg hier je Canadese velden toe zoals nodig
]

# -------------------------------------------------------
# Excel laden
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
# Stieren sorteren
# -------------------------------------------------------
def custom_sort_ras(df):
    if "Ras" not in df.columns:
        df["Ras"] = ""
    if "naam" not in df.columns:
        df["naam"] = ""
    order_map = {"Holstein zwartbont": 1, "Red Holstein": 2}
    df["ras_sort"] = df["Ras"].map(order_map).fillna(3)
    df_sorted = df.sort_values(by=["ras_sort", "naam"], ascending=True)
    df_sorted.drop(columns=["ras_sort"], inplace=True)
    return df_sorted

# -------------------------------------------------------
# Top 5-tabellen maken
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
    if df.empty:
        return pd.DataFrame()

    df["Ras_clean"] = df["Ras"].astype(str).str.strip().str.lower()
    df = df[df["Ras_clean"].isin([
        "holstein zwartbont",
        "holstein zwartbont + rf",
        "red holstein"
    ])].copy()

    for fok in fokwaarden:
        if fok not in df.columns:
            df[fok] = pd.NA

        block = []
        header_row = {
            "Fokwaarde": fok,
            "zwartbont_stier": "Stier",
            "zwartbont_value": "Waarde",
            "roodbont_stier": "Stier",
            "roodbont_value": "Waarde"
        }
        block.append(header_row)

        df_z = df[df["Ras_clean"].isin(["holstein zwartbont", "holstein zwartbont + rf"])].copy()
        df_z[fok] = pd.to_numeric(df_z[fok], errors='coerce')
        df_z = df_z.sort_values(by=fok, ascending=False)

        df_r = df[df["Ras_clean"].str.contains("red holstein")].copy()
        df_r[fok] = pd.to_numeric(df_r[fok], errors='coerce')
        df_r = df_r.sort_values(by=fok, ascending=False)

        for i in range(5):
            row = {
                "Fokwaarde": "",
                "zwartbont_stier": "",
                "zwartbont_value": "",
                "roodbont_stier": "",
                "roodbont_value": ""
            }
            if i < len(df_z):
                row["zwartbont_stier"] = str(df_z.iloc[i]["naam"])
                row["zwartbont_value"] = str(df_z.iloc[i][fok])
            if i < len(df_r):
                row["roodbont_stier"] = str(df_r.iloc[i]["naam"])
                row["roodbont_value"] = str(df_r.iloc[i][fok])
            block.append(row)

        block.append({
            "Fokwaarde": "",
            "zwartbont_stier": "",
            "zwartbont_value": "",
            "roodbont_stier": "",
            "roodbont_value": ""
        })

        blocks.extend(block)

    return pd.DataFrame(blocks)

# -------------------------------------------------------
# Main
# -------------------------------------------------------
def main():
    st.set_page_config(layout="wide")
    st.title("Stierenkaart Generator")

    # Taalkeuze
    taal = st.selectbox("Kies stierenkaart type:", ["Nederland", "Canada"])

    mapping_table = mapping_table_nl if taal == "Nederland" else mapping_table_canada

    uploaded_file = st.file_uploader("Upload Excel bestand", type=["xlsx"])
    if uploaded_file:
        df_raw = load_excel(uploaded_file)

        if df_raw is not None:
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

            # pinkenstier berekenen (alleen NL)
            if taal == "Nederland" and "geboortegemak" in df_mapped.columns:
                df_mapped["pinkenstier"] = df_mapped["geboortegemak"].apply(
                    lambda x: "p" if pd.notna(x) and x > 100 else ""
                )
            else:
                df_mapped["pinkenstier"] = ""

            # Kolomvolgorde
            kolomvolgorde = ["superbevruchter","ki-code","naam","pinkenstier"] + \
                [x["Stierenkaart"] for x in mapping_table if x["Stierenkaart"] not in ["superbevruchter","ki-code","naam"]]
            kolomvolgorde = [x for x in kolomvolgorde if x in df_mapped.columns]

            overige_kolommen = [c for c in df_mapped.columns if c not in kolomvolgorde]
            df_mapped = df_mapped[kolomvolgorde + overige_kolommen]

            # Voeg Display toe
            if "ki-code" in df_mapped.columns and "naam" in df_mapped.columns:
                df_mapped["Display"] = df_mapped["ki-code"].astype(str) + " - " + df_mapped["naam"].astype(str)

                selected_display = st.multiselect("Selecteer stieren:", df_mapped["Display"].tolist())
                if selected_display:
                    selected_codes = [x.split(" - ")[0] for x in selected_display]
                    df_selected = df_mapped[df_mapped["ki-code"].isin(selected_codes)].copy()
                    df_selected = custom_sort_ras(df_selected)

                    st.dataframe(df_selected)

                    # Top 5 tabellen
                    df_top5 = create_top5_table(df_selected)
                    if not df_top5.empty:
                        st.subheader("Top 5 tabellen")
                        st.dataframe(df_top5)

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_selected.to_excel(writer, sheet_name='Stierenkaart', index=False)
                        if not df_top5.empty:
                            df_top5.to_excel(writer, sheet_name='Top5_per_ras', index=False)

                    st.download_button("Download Excel", output.getvalue(), file_name="stierenkaart.xlsx")
                else:
                    st.info("Selecteer stieren om te downloaden.")
            else:
                st.warning("Kolommen 'ki-code' en/of 'naam' ontbreken.")
        else:
            st.error("Fout bij inlezen bestand.")

if __name__ == "__main__":
    main()
