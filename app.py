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
    {"Stierenkaart": "prijs", "Titel in bestand": "Prijs", "Formule": None},  # Wordt overschreven met prijslijst
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
    {"Stierenkaart": "persistentie", "Titel in bestand": "", "Formule": "/100"},
    {"Stierenkaart": "klauwgezondheid", "Titel in bestand": "OFFICIAL CLAW HEALTH EVALUATION IN THIS COUNTRY klauwgezondheid", "Formule": "/100"},
    {"Stierenkaart": "levensduur", "Titel in bestand": "OFFICIAL CALF LIVABILITY EVALUATION IN THIS COUNTRY levensduur", "Formule": "/100"},
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
# Prijslijst inlezen (normaal en gesekst) - FLEXIBEL
# -------------------------------------------------------
def load_prijslijst(file):
    try:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()

        if "ki-code" not in df.columns:
            st.error("Kolom 'ki-code' ontbreekt in de prijslijst.")
            return None, None

        # Kies prijs-kolom in volgorde van voorkeur
        prijs_kandidaten = ["Nieuwe prijs", "Prijs op stierenkaart", "Eenheidsprijs", "prijs"]
        prijs_kolom = next((c for c in prijs_kandidaten if c in df.columns), None)
        if not prijs_kolom:
            st.error("Geen bruikbare prijs-kolom gevonden (gezocht: 'Nieuwe prijs', 'Prijs op stierenkaart', 'Eenheidsprijs', 'prijs').")
            return None, None

        df = df.rename(columns={prijs_kolom: "prijs"})
        df["ki-code"] = df["ki-code"].astype(str).str.strip().str.upper()
        df["prijs"] = pd.to_numeric(df["prijs"].astype(str).str.replace(",", ".", regex=False), errors="coerce")

        # -S betekent gesekst; voor de merge gebruiken we de basiscode
        df["is_gesekst"] = df["ki-code"].str.endswith("-S")
        df_normaal = df[~df["is_gesekst"]][["ki-code", "prijs"]].copy()
        df_gesekst = df[df["is_gesekst"]][["ki-code", "prijs"]].copy()
        df_gesekst["ki-code"] = df_gesekst["ki-code"].str.replace("-S", "", regex=False)
        df_gesekst = df_gesekst.rename(columns={"prijs": "prijs gesekst"})

        return df_normaal, df_gesekst
    except Exception as e:
        st.error(f"Fout bij laden prijslijst: {e}")
        return None, None

# -------------------------------------------------------
# Stieren sorteren
# -------------------------------------------------------
def custom_sort_ras(df, df_raw=None):
    if "Ras" not in df.columns:
        # Probeer Rasomschrijving vanuit het bronbestand te vullen
        if df_raw is not None and "Rasomschrijving" in df_raw.columns:
            df["Ras"] = df_raw["Rasomschrijving"].fillna("")
        else:
            df["Ras"] = ""
    if "naam" not in df.columns:
        df["naam"] = ""
    order_map = {"Holstein zwartbont": 1, "Red Holstein": 2}
    df["ras_sort"] = df["Ras"].map(order_map).fillna(3)
    df_sorted = df.sort_values(by=["ras_sort", "naam"], ascending=True)
    return df_sorted.drop(columns=["ras_sort"])

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

    st.success(f"PIM ingelezen met {len(df_raw)} rijen en {len(df_raw.columns)} kolommen.")
    if st.checkbox("Toon kolomnamen PIM"):
        st.write(df_raw.columns.tolist())

    # --- Mapping toepassen ---
    final_data = {}
    for mapping in mapping_table_pim:
        titel = mapping["Titel in bestand"]
        std_naam = mapping["Stierenkaart"]
        formule = mapping["Formule"]
        if titel and titel in df_raw.columns:
            # Veilige replace (voorkomt rare type-issues)
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

    # KI-codes normaliseren voor merge en display
    if "ki-code" in df_mapped.columns:
        df_mapped["ki-code"] = (
            df_mapped["ki-code"]
            .astype(str).str.strip().str.upper()
            .str.replace(r"\.0$", "", regex=True)
        )

    # --- Prijslijst verwerken (indien aanwezig) ---
    if prijs_file:
        df_prijs_normaal, df_prijs_gesekst = load_prijslijst(prijs_file)
        if df_prijs_normaal is not None:
            # haal bestaande prijs-kolommen weg om schone merge te krijgen
            df_mapped = df_mapped.drop(columns=["prijs", "prijs gesekst"], errors="ignore")
            df_mapped = df_mapped.merge(df_prijs_normaal, on="ki-code", how="left")
            df_mapped = df_mapped.merge(df_prijs_gesekst, on="ki-code", how="left")
            st.success("Normale en gesekste prijzen succesvol bijgewerkt vanuit prijslijst.")
        else:
            st.warning("Kon prijslijst niet verwerken.")
    else:
        st.warning("Geen prijslijst geüpload. Oorspronkelijke prijswaarden uit PIM worden gebruikt (indien aanwezig).")

    # Pinkenstier label
    if "geboortegemak" in df_mapped.columns:
        df_mapped["pinkenstier"] = df_mapped["geboortegemak"].apply(
            lambda x: "p" if pd.notna(x) and x > 100 else ""
        )
    else:
        df_mapped["pinkenstier"] = ""

    # Display-veld
    if "ki-code" in df_mapped.columns and "naam" in df_mapped.columns:
        df_mapped["Display"] = df_mapped["ki-code"].astype(str) + " - " + df_mapped["naam"].astype(str)
        selected_display = st.multiselect("Selecteer stieren:", options=df_mapped["Display"].tolist())

        # Debug-blok (optioneel aan te zetten)
        if st.checkbox("Toon debug- en match-informatie"):
            st.write("Aantal stieren:", len(df_mapped))
            st.write("Unieke KI-codes:", df_mapped["ki-code"].nunique())
            if prijs_file and df_prijs_normaal is not None:
                st.write("Aantal met prijs:", int(df_mapped["prijs"].notna().sum()))
                st.write("Aantal met prijs gesekst:", int(df_mapped["prijs gesekst"].notna().sum()))

        if selected_display:
            selected_codes = [x.split(" - ")[0] for x in selected_display]
            df_selected = df_mapped[df_mapped["ki-code"].isin(selected_codes)].copy()
            df_selected = custom_sort_ras(df_selected, df_raw=df_raw)

            st.subheader("Geselecteerde stieren")
            st.dataframe(df_selected, use_container_width=True)

            # Exporteren
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_selected.to_excel(writer, sheet_name='Stierenkaart', index=False)

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
