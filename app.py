import os
os.environ["STREAMLIT_WATCH"] = "false"

import streamlit as st
import pandas as pd
import io

# -------------------------------------------------------
# Mapping table PIM (met formules)
# -------------------------------------------------------
mapping_table_pim = [
    {"Stierenkaart": "superbevruchter", "Titel in bestand": "Superbevruchter", "Formule": None},
    {"Stierenkaart": "ki-code", "Titel in bestand": "Stiercode NL / KI code", "Formule": None},
    {"Stierenkaart": "naam", "Titel in bestand": "Afkorting stier (zoeknaam)", "Formule": None},
    {"Stierenkaart": "vader", "Titel in bestand": "Roepnaam Vader", "Formule": None},
    {"Stierenkaart": "vaders vader", "Titel in bestand": "Roepnaam Moeders Vader", "Formule": None},
    {"Stierenkaart": "PFW", "Titel in bestand": "PFW code", "Formule": None},
    {"Stierenkaart": "aAa", "Titel in bestand": "AAa code", "Formule": None},
    {"Stierenkaart": "Beta caseine", "Titel in bestand": "Betacasine", "Formule": None},
    {"Stierenkaart": "Kappa caseine", "Titel in bestand": "Kappa-caseine", "Formule": None},
    {"Stierenkaart": "prijs", "Titel in bestand": "Prijs", "Formule": None},  # wordt overschreven met prijslijst
    {"Stierenkaart": "prijs gesekst", "Titel in bestand": "", "Formule": None},
    {"Stierenkaart": "&betrouwbaarheid productie", "Titel in bestand": "Official Production Evaluation in this Country %betrouwbaarheid (Productie-index)", "Formule": None},
    {"Stierenkaart": "kg melk", "Titel in bestand": "Official Production Evalution in this Country KG Melk", "Formule": "/10"},
    {"Stierenkaart": "%vet", "Titel in bestand": "Offical Production Evaluation in this Country %vet", "Formule": "/100"},
    {"Stierenkaart": "%eiwit", "Titel in bestand": "Official Production Evaluation in this County %eiwit", "Formule": "/100"},
    {"Stierenkaart": "kg vet", "Titel in bestand": "Official Production Evaluation in this Country KG vet", "Formule": "/10"},
    {"Stierenkaart": "kg eiwit", "Titel in bestand": "Official Production Evaluation in this Country KG eiwit", "Formule": "/10"},
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
    {"Stierenkaart": "persistentie", "Titel in bestand": "Persistentie", "Formule": "/100"},
    {"Stierenkaart": "klauwgezondheid", "Titel in bestand": "OFFICIAL CLAW HEALTH EVALUATION IN THIS COUNTRY klauwgezondheid", "Formule": "/100"},
    {"Stierenkaart": "levensduur", "Titel in bestand": "OFFICIAL CALF LIVABILITY EVALUATION IN THIS COUNTRY levensduur", "Formule": None},
    {"Stierenkaart": "koe familie", "Titel in bestand": "Koe familie", "Formule": None}

]

# -------------------------------------------------------
# Excel inlezen
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
# Prijslijst inlezen - ALTIJD prijs uit kolom E (5e kolom)
# -------------------------------------------------------
def load_prijslijst(file):
    """
    Neemt prijs altijd uit kolom E (ongeacht kolomtitel).
    Splitst -S (gesekst) naar aparte kolom 'prijs gesekst'.
    """
    try:
        df = pd.read_excel(file)  # evt. sheet_name="Blad1" toevoegen indien gewenst
        df.columns = df.columns.str.strip()

        if df.shape[1] < 5:
            st.error("De prijslijst heeft minder dan 5 kolommen; kolom E ontbreekt.")
            return None, None

        # kolom A = ki-code (eerste kolom), kolom E = prijs (vijfde kolom)
        ki_series = df.iloc[:, 0]
        prijs_series = df.iloc[:, 4]

        df_prices = pd.DataFrame({
            "ki-code": ki_series.astype(str).str.strip().str.upper(),
            "prijs": prijs_series.astype(str).str.replace(",", ".", regex=False),
        })
        df_prices["prijs"] = pd.to_numeric(df_prices["prijs"], errors="coerce")

        # -S betekent gesekst; verwijder -S voor join-sleutel
        df_prices["is_gesekst"] = df_prices["ki-code"].str.endswith("-S")
        df_normaal = df_prices[~df_prices["is_gesekst"]][["ki-code", "prijs"]].copy()
        df_gesekst = df_prices[df_prices["is_gesekst"]][["ki-code", "prijs"]].copy()
        df_gesekst["ki-code"] = df_gesekst["ki-code"].str.replace("-S", "", regex=False)
        df_gesekst = df_gesekst.rename(columns={"prijs": "prijs gesekst"})

        return df_normaal, df_gesekst
    except Exception as e:
        st.error(f"Fout bij laden prijslijst (kolom E): {e}")
        return None, None

# -------------------------------------------------------
# Stieren sorteren
# -------------------------------------------------------
def custom_sort_ras(df, df_raw=None):
    if "Ras" not in df.columns and df_raw is not None and "Rasomschrijving" in df_raw.columns:
        df["Ras"] = df_raw["Rasomschrijving"].fillna("")
    elif "Ras" not in df.columns:
        df["Ras"] = ""
    if "naam" not in df.columns:
        df["naam"] = ""
    order_map = {"Holstein zwartbont": 1, "Red Holstein": 2}
    df["ras_sort"] = df["Ras"].map(order_map).fillna(3)
    return df.sort_values(by=["ras_sort", "naam"], ascending=True).drop(columns=["ras_sort"])

# -------------------------------------------------------
# Streamlit main
# -------------------------------------------------------
def main():
    st.set_page_config(layout="wide")
    st.title("Stierenkaart Generator (PIM versie, met formules)")

    uploaded_file = st.file_uploader("Upload PIM K.I. Samen.xlsx", type=["xlsx"])
    prijs_file = st.file_uploader("Upload prijslijst.xlsx", type=["xlsx"])

    if not uploaded_file:
        st.info("Upload eerst het PIM-bestand.")
        return

    df_raw = load_excel(uploaded_file)
    if df_raw is None:
        st.error("Kon het PIM-bestand niet inlezen.")
        return

    # --- Mapping toepassen ---
    final_data = {}
    for mapping in mapping_table_pim:
        titel = mapping["Titel in bestand"]
        std_naam = mapping["Stierenkaart"]
        formule = mapping["Formule"]
        if titel and titel in df_raw.columns:
            kolom = df_raw[titel].replace({99999: pd.NA, "+999": pd.NA})
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

    # KI-codes normaliseren (ook voor merge)
    if "ki-code" in df_mapped.columns:
        df_mapped["ki-code"] = (
            df_mapped["ki-code"]
            .astype(str).str.strip().str.upper()
            .str.replace(r"\.0$", "", regex=True)
        )

    # --- Prijslijst verwerken ---
    if prijs_file:
        df_prijs_normaal, df_prijs_gesekst = load_prijslijst(prijs_file)
        if df_prijs_normaal is not None:
            df_mapped = df_mapped.drop(columns=["prijs", "prijs gesekst"], errors="ignore")
            df_mapped = df_mapped.merge(df_prijs_normaal, on="ki-code", how="left")
            df_mapped = df_mapped.merge(df_prijs_gesekst, on="ki-code", how="left")

    # Pinkenstier
    if "geboortegemak" in df_mapped.columns:
        df_mapped["pinkenstier"] = df_mapped["geboortegemak"].apply(
            lambda x: "p" if pd.notna(x) and x > 100 else ""
        )
    else:
        df_mapped["pinkenstier"] = ""

    # --- UI selectie & export ---
    if "ki-code" in df_mapped.columns and "naam" in df_mapped.columns:
        df_mapped["Display"] = df_mapped["ki-code"] + " - " + df_mapped["naam"]
        selected_display = st.multiselect("Selecteer stieren:", options=df_mapped["Display"].tolist())

        if selected_display:
            selected_codes = [x.split(" - ")[0] for x in selected_display]
            df_selected = df_mapped[df_mapped["ki-code"].isin(selected_codes)].copy()
            df_selected = custom_sort_ras(df_selected, df_raw=df_raw)

            st.subheader("Geselecteerde stieren")
            st.dataframe(df_selected, use_container_width=True)

            # -------------------------------------------------------
            # Exporteren met gewenste kolomvolgorde
            # -------------------------------------------------------
            output = io.BytesIO()

            # maak lege TIP-kolom als die ontbreekt
            if "TIP" not in df_selected.columns:
                df_selected["TIP"] = ""

            kolomvolgorde = [
                "superbevruchter", "ki-code", "naam", "pinkenstier", "vader", "vaders vader",
                "PFW", "aAa", "Beta caseine", "Kappa caseine",
                "prijs", "prijs gesekst",
                "&betrouwbaarheid productie", "kg melk", "%vet", "%eiwit", "kg vet", "kg eiwit",
                "INET", "NVI", "TIP",
                "%betrouwbaarheid exterieur", "frame", "uier", "benen", "totaal",
                "hoogtemaat", "voorhand", "inhoud", "ribvorm", "conditiescore",
                "kruisligging", "kruisbreedte", "beenstand achter", "beenstand zij",
                "klauwhoek", "voorbeenstand", "beengebruik",
                "vooruieraanhechting", "voorspeenplaatsing", "speenlengte",
                "uierdiepte", "achteruierhoogte", "ophangband",
                "achterspeenplaatsing",
                "geboortegemak", "melksnelheid", "celgetal", "vruchtbaarheid",
                "karakter", "laatrijpheid", "persistentie", "klauwgezondheid", "levensduur", "koe familie"
            ]

            # Alleen kolommen meenemen die er zijn
            bestaande_kolommen = [c for c in kolomvolgorde if c in df_selected.columns]

            # Prijzen afronden
            for c in ["prijs", "prijs gesekst"]:
                if c in df_selected.columns:
                    df_selected[c] = pd.to_numeric(df_selected[c], errors="coerce").round(2)

            df_selected_export = df_selected[bestaande_kolommen].copy()

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_selected_export.to_excel(writer, sheet_name='Stierenkaart', index=False)

            st.download_button(
                label="Download selectie",
                data=output.getvalue(),
                file_name="stierenkaart_selectie.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Selecteer één of meer stieren om de gegevens te zien en te downloaden.")
    else:
        st.warning("Kolommen 'ki-code' en/of 'naam' ontbreken in de gemapte data.")

if __name__ == "__main__":
    main()
