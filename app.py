import os
os.environ["STREAMLIT_WATCH"] = "false"

import streamlit as st
import pandas as pd
import io

# -------------------------------------------------------
# Mapping table PIM
# -------------------------------------------------------
mapping_table_pim = [
    {"Stierenkaart": "superbevruchter", "Titel in bestand": "Superbevruchter"},
    {"Stierenkaart": "ki-code", "Titel in bestand": "Stiercode NL / KI code"},
    {"Stierenkaart": "naam", "Titel in bestand": "Afkorting stier (zoeknaam)"},
    {"Stierenkaart": "pinkenstier", "Titel in bestand": "p wanneer geboortegemak > 100"},
    {"Stierenkaart": "vader", "Titel in bestand": "Roepnaam Vader"},
    {"Stierenkaart": "vaders vader", "Titel in bestand": "Roepnaam Vaders Vader"},
    {"Stierenkaart": "PFW", "Titel in bestand": "PFW code"},
    {"Stierenkaart": "aAa", "Titel in bestand": "AAa code"},
    {"Stierenkaart": "Beta caseine", "Titel in bestand": "Betacasine"},
    {"Stierenkaart": "Kappa caseine", "Titel in bestand": "Kappa-caseine"},
    {"Stierenkaart": "prijs", "Titel in bestand": "Prijs"},
    {"Stierenkaart": "prijs gesekst", "Titel in bestand": ""},
    {"Stierenkaart": "&betrouwbaarheid productie", "Titel in bestand": "Official Production Evaluation in this Country %betrouwbaarheid (Productie-index)"},
    {"Stierenkaart": "kg melk", "Titel in bestand": "Official Production Evalution in this Country KG Melk"},
    {"Stierenkaart": "%vet", "Titel in bestand": "Offical Production Evaluation in this Country %vet"},
    {"Stierenkaart": "%eiwit", "Titel in bestand": "Official Production Evaluation in this County %eiwit"},
    {"Stierenkaart": "kg vet", "Titel in bestand": "Official Production Evaluation in this Country KG vet"},
    {"Stierenkaart": "kg eiwit", "Titel in bestand": "Official Production Evaluation in this Country KG eiwit"},
    {"Stierenkaart": "INET", "Titel in bestand": "Official Production Evaluation in this Country Inet"},
    {"Stierenkaart": "NVI", "Titel in bestand": "Official Production Evaluation in this Country NVI"},
    {"Stierenkaart": "TIP", "Titel in bestand": ""},
    {"Stierenkaart": "%betrouwbaarheid exterieur", "Titel in bestand": "%betrouwbaarheid (exterieur-index)"},
    {"Stierenkaart": "frame", "Titel in bestand": "GENERAL CHARACTERISTICS frame"},
    {"Stierenkaart": "uier", "Titel in bestand": "GENERAL CHARACTERISTICS uier"},
    {"Stierenkaart": "benen", "Titel in bestand": "GENERAL CHARACTERISTICS benen"},
    {"Stierenkaart": "totaal", "Titel in bestand": "GENERAL CHARACTERISTICS totaal (Exterieur-index)"},
    {"Stierenkaart": "hoogtemaat", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY hoogtemaat"},
    {"Stierenkaart": "voorhand", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorhand"},
    {"Stierenkaart": "inhoud", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY inhoud"},
    {"Stierenkaart": "ribvorm", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY openheid"},
    {"Stierenkaart": "conditiescore", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY conditiescore"},
    {"Stierenkaart": "kruisligging", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY kruisligging"},
    {"Stierenkaart": "kruisbreedte", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY kruisbreedte"},
    {"Stierenkaart": "beenstand achter", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achteraanzicht benen"},
    {"Stierenkaart": "beenstand zij", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY beenstand zij"},
    {"Stierenkaart": "klauwhoek", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY klauwhoek"},
    {"Stierenkaart": "voorbeenstand", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorbeenstand"},
    {"Stierenkaart": "beengebruik", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY beengebruik"},
    {"Stierenkaart": "vooruieraanhechting", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY vooruieraanhechting"},
    {"Stierenkaart": "voorspeenplaatsing", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorspeenplaatsing"},
    {"Stierenkaart": "speenlengte", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY speenlengte"},
    {"Stierenkaart": "uierdiepte", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY uierdiepte"},
    {"Stierenkaart": "achteruierhoogte", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achteruierhoogte"},
    {"Stierenkaart": "ophangband", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY ophangband"},
    {"Stierenkaart": "achterspeenplaatsing", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achterspeenplaatsing"},
    {"Stierenkaart": "geboortegemak", "Titel in bestand": "OFFICIAL CALVING EASE EVALUATION IN THIS COUNTRY geboortegemak"},
    {"Stierenkaart": "melksnelheid", "Titel in bestand": "OFFICIAL MILKING SPEED AND TEMPERAMENT EVALUATION IN THIS COUNTRY melksnelheid"},
    {"Stierenkaart": "celgetal", "Titel in bestand": "OFFICIAL SOMATIC CELL COUNT EVALUATION IN THIS COUNTRY celgetal"},
    {"Stierenkaart": "vruchtbaarheid", "Titel in bestand": "OFFICIAL FEMALE FERTILITY EVALUATION IN THIS COUNTRY vruchtbaarheid"},
    {"Stierenkaart": "karakter", "Titel in bestand": "OFFICIAL MILKING SPEED AND TEMPERAMENT EVALUATION IN THIS COUNTRY karakter"},
    {"Stierenkaart": "laatrijpheid", "Titel in bestand": "OFFICIAL CALVING EASE EVALUATION IN THIS COUNTRY laatrijpheid"},
    {"Stierenkaart": "persistentie", "Titel in bestand": ""},
    {"Stierenkaart": "klauwgezondheid", "Titel in bestand": "OFFICIAL CLAW HEALTH EVALUATION IN THIS COUNTRY klauwgezondheid"},
    {"Stierenkaart": "levensduur", "Titel in bestand": "OFFICIAL CALF LIVABILITY EVALUATION IN THIS COUNTRY levensduur"}
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
# Streamlit main
# -------------------------------------------------------
def main():
    st.set_page_config(layout="wide")
    st.title("Stierenkaart Generator (PIM versie)")

    uploaded_file = st.file_uploader("Upload PIM K.I. Samen.xlsx", type=["xlsx"])

    if uploaded_file:
        df_raw = load_excel(uploaded_file)

        if df_raw is not None:
            st.success(f"Bestand ingelezen met {len(df_raw)} rijen en {len(df_raw.columns)} kolommen.")

            if st.checkbox("Toon kolomnamen"):
                st.write(df_raw.columns.tolist())

            # Maak mapping
            final_data = {}
            for mapping in mapping_table_pim:
                titel = mapping["Titel in bestand"]
                std_naam = mapping["Stierenkaart"]
                if titel and titel in df_raw.columns:
                    final_data[std_naam] = df_raw[titel]
                else:
                    final_data[std_naam] = ""

            df_mapped = pd.DataFrame(final_data)

            # Bouw display kolom
            if "ki-code" in df_mapped.columns and "naam" in df_mapped.columns:
                df_mapped["ki-code"] = df_mapped["ki-code"].astype(str).str.strip().str.upper()
                df_mapped["Display"] = df_mapped["ki-code"] + " - " + df_mapped["naam"].astype(str)

                selected_display = st.multiselect(
                    "Selecteer stieren:",
                    options=df_mapped["Display"].tolist()
                )

                if selected_display:
                    selected_codes = [x.split(" - ")[0] for x in selected_display]
                    df_selected = df_mapped[df_mapped["ki-code"].isin(selected_codes)].copy()

                    df_selected = custom_sort_ras(df_selected)

                    st.subheader("Geselecteerde stieren")
                    st.dataframe(df_selected, use_container_width=True)

                    df_top5 = create_top5_table(df_selected)
                    if not df_top5.empty:
                        st.subheader("Top 5-tabellen per fokwaarde")
                        st.dataframe(df_top5, use_container_width=True)

                    # Excel download
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_selected.to_excel(writer, sheet_name='Stierenkaart', index=False)
                        if not df_top5.empty:
                            df_top5.to_excel(writer, sheet_name='Top5_per_ras', index=False)

                    st.download_button(
                        label="Download selectie + Top 5-tabellen",
                        data=output.getvalue(),
                        file_name="stierenkaart_selectie.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info("Selecteer één of meer stieren om de gegevens te zien en te downloaden.")
            else:
                st.warning("Kolommen 'ki-code' en/of 'naam' ontbreken in de gemapte data.")
        else:
            st.error("Kon het bestand niet inlezen.")
    else:
        st.info("Upload eerst het PIM-bestand.")

if __name__ == "__main__":
    main()
