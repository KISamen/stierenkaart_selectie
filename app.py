import os
os.environ["STREAMLIT_WATCH"] = "false"

import streamlit as st
import pandas as pd
import io

# -------------------------------------------------------
# Mapping table PIM (Nederland)
# -------------------------------------------------------
mapping_table_pim = [
    {"Stierenkaart": "superbevruchter",                   "Titel in bestand": "Superbevruchter",                                                                              "Formule": None},
    {"Stierenkaart": "ki-code",                           "Titel in bestand": "Stiercode NL / KI code",                                                                      "Formule": None},
    {"Stierenkaart": "naam",                              "Titel in bestand": "Afkorting stier (zoeknaam)",                                                                   "Formule": None},
    {"Stierenkaart": "vader",                             "Titel in bestand": "Roepnaam Vader",                                                                               "Formule": None},
    {"Stierenkaart": "vaders vader",                      "Titel in bestand": "Roepnaam Moeders Vader",                                                                       "Formule": None},
    {"Stierenkaart": "PFW",                               "Titel in bestand": "PFW code",                                                                                     "Formule": None},
    {"Stierenkaart": "aAa",                               "Titel in bestand": "AAa code",                                                                                     "Formule": None},
    {"Stierenkaart": "Beta caseine",                      "Titel in bestand": "Betacasine",                                                                                   "Formule": None},
    {"Stierenkaart": "Kappa caseine",                     "Titel in bestand": "Kappa-caseine",                                                                                "Formule": None},
    {"Stierenkaart": "prijs",                             "Titel in bestand": "Prijs",                                                                                       "Formule": None},
    {"Stierenkaart": "prijs gesekst",                     "Titel in bestand": "",                                                                                             "Formule": None},
    {"Stierenkaart": "&betrouwbaarheid productie",        "Titel in bestand": "Official Production Evaluation in this Country %betrouwbaarheid (Productie-index)",            "Formule": None},
    {"Stierenkaart": "kg melk",                           "Titel in bestand": "Official Production Evalution in this Country KG Melk",                                         "Formule": "/10"},
    {"Stierenkaart": "%vet",                              "Titel in bestand": "Offical Production Evaluation in this Country %vet",                                           "Formule": "/100"},
    {"Stierenkaart": "%eiwit",                            "Titel in bestand": "Official Production Evaluation in this County %eiwit",                                        "Formule": "/100"},
    {"Stierenkaart": "kg vet",                            "Titel in bestand": "Official Production Evaluation in this Country KG vet",                                         "Formule": "/10"},
    {"Stierenkaart": "kg eiwit",                          "Titel in bestand": "Official Production Evaluation in this Country KG eiwit",                                       "Formule": None},
    {"Stierenkaart": "INET",                              "Titel in bestand": "Official Production Evaluation in this Country Inet",                                           "Formule": None},
    {"Stierenkaart": "NVI",                               "Titel in bestand": "Official Production Evaluation in this Country NVI",                                            "Formule": None},
    {"Stierenkaart": "TIP",                               "Titel in bestand": "",                                                                                             "Formule": None},
    {"Stierenkaart": "%betrouwbaarheid exterieur",        "Titel in bestand": "%betrouwbaarheid (exterieur-index)",                                                           "Formule": None},
    {"Stierenkaart": "frame",                             "Titel in bestand": "GENERAL CHARACTERISTICS frame",                                                               "Formule": "/100"},
    {"Stierenkaart": "uier",                              "Titel in bestand": "GENERAL CHARACTERISTICS uier",                                                                "Formule": "/100"},
    {"Stierenkaart": "benen",                             "Titel in bestand": "GENERAL CHARACTERISTICS benen",                                                              "Formule": "/100"},
    {"Stierenkaart": "totaal",                            "Titel in bestand": "GENERAL CHARACTERISTICS totaal (Exterieur-index)",                                            "Formule": "/100"},
    {"Stierenkaart": "hoogtemaat",                        "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY hoogtemaat",                               "Formule": "/100"},
    {"Stierenkaart": "voorhand",                          "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorhand",                                "Formule": "/100"},
    {"Stierenkaart": "inhoud",                            "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY inhoud",                                  "Formule": "/100"},
    {"Stierenkaart": "ribvorm",                           "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY openheid",                               "Formule": "/100"},
    {"Stierenkaart": "conditiescore",                     "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY conditiescore",                         "Formule": "/100"},
    {"Stierenkaart": "kruisligging",                      "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY kruisligging",                          "Formule": "/100"},
    {"Stierenkaart": "kruisbreedte",                      "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY kruisbreedte",                          "Formule": "/100"},
    {"Stierenkaart": "beenstand achter",                  "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achteraanzicht benen",                "Formule": "/100"},
    {"Stierenkaart": "beenstand zij",                     "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY beenstand zij",                       "Formule": "/100"},
    {"Stierenkaart": "klauwhoek",                         "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY klauwhoek",                         "Formule": "/100"},
    {"Stierenkaart": "voorbeenstand",                     "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorbeenstand",                     "Formule": "/100"},
    {"Stierenkaart": "beengebruik",                       "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY beengebruik",                       "Formule": "/100"},
    {"Stierenkaart": "vooruieraanhechting",               "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY vooruieraanhechting",             "Formule": "/100"},
    {"Stierenkaart": "voorspeenplaatsing",                "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorspeenplaatsing",            "Formule": "/100"},
    {"Stierenkaart": "speenlengte",                       "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY speenlengte",                    "Formule": "/100"},
    {"Stierenkaart": "uierdiepte",                        "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY uierdiepte",                    "Formule": "/100"},
    {"Stierenkaart": "achteruierhoogte",                  "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achteruierhoogte",              "Formule": "/100"},
    {"Stierenkaart": "ophangband",                        "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY ophangband",                  "Formule": "/100"},
    {"Stierenkaart": "achterspeenplaatsing",              "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achterspeenplaatsing",        "Formule": "/100"},
    {"Stierenkaart": "geboortegemak",                     "Titel in bestand": "OFFICIAL CALVING EASE EVALUATION IN THIS COUNTRY geboortegemak",            "Formule": "/100"},
    {"Stierenkaart": "melksnelheid",                      "Titel in bestand": "OFFICIAL MILKING SPEED AND TEMPERAMENT EVALUATION IN THIS COUNTRY melksnelheid","Formule": "/100"},
    {"Stierenkaart": "celgetal",                          "Titel in bestand": "OFFICIAL SOMATIC CELL COUNT EVALUATION IN THIS COUNTRY celgetal",          "Formule": "/100"},
    {"Stierenkaart": "vruchtbaarheid",                    "Titel in bestand": "OFFICIAL FEMALE FERTILITY EVALUATION IN THIS COUNTRY vruchtbaarheid",     "Formule": "/100"},
    {"Stierenkaart": "karakter",                          "Titel in bestand": "OFFICIAL MILKING SPEED AND TEMPERAMENT EVALUATION IN THIS COUNTRY karakter","Formule": "/100"},
    {"Stierenkaart": "laatrijpheid",                      "Titel in bestand": "OFFICIAL CALVING EASE EVALUATION IN THIS COUNTRY laatrijpheid",           "Formule": "/100"},
    {"Stierenkaart": "persistentie",                      "Titel in bestand": "",                                                                                       "Formule": "/100"},
    {"Stierenkaart": "klauwgezondheid",                   "Titel in bestand": "OFFICIAL CLAW HEALTH EVALUATION IN THIS COUNTRY klauwgezondheid",   "Formule": "/100"},
    {"Stierenkaart": "levensduur",                        "Titel in bestand": "OFFICIAL CALF LIVABILITY EVALUATION IN THIS COUNTRY levensduur",   "Formule": "/100"}
]

# -------------------------------------------------------
# Mapping table PIM (Canada)
# -------------------------------------------------------
mapping_table_ca = [
    {"Stierenkaart": "ki-code",                               "Titel in bestand": "Stiercode NL / KI code",                                                                     "Formule": None},
    {"Stierenkaart": "Name",                                  "Titel in bestand": "Afkorting stier (zoeknaam)",                                                                 "Formule": None},
    {"Stierenkaart": "Pedigree father",                       "Titel in bestand": "Roepnaam Vader",                                                                             "Formule": None},
    {"Stierenkaart": "Pedigree grandfather",                  "Titel in bestand": "Roepnaam Vaders Vader",                                                                     "Formule": None},
    {"Stierenkaart": "aAa",                                   "Titel in bestand": "AAa code",                                                                                    "Formule": None},
    {"Stierenkaart": "prijs",                                 "Titel in bestand": "Prijs",                                                                                      "Formule": None},
    {"Stierenkaart": "prijs gesekst",                         "Titel in bestand": "",                                                                                          "Formule": None},
    {"Stierenkaart": "%reliability",                          "Titel in bestand": "Official Production Evaluation in this Country %betrouwbaarheid (Productie-index)",         "Formule": None},
    {"Stierenkaart": "kg milk",                               "Titel in bestand": "Official Production Evalution in this Country KG Melk",                                      "Formule": "/10"},
    {"Stierenkaart": "%fat",                                  "Titel in bestand": "Offical Production Evaluation in this Country %vet",                                        "Formule": "/100"},
    {"Stierenkaart": "%protein",                              "Titel in bestand": "Official Production Evaluation in this County %eiwit",                                     "Formule": "/100"},
    {"Stierenkaart": "kg fat",                                "Titel in bestand": "Official Production Evaluation in this Country KG vet",                                     "Formule": "/10"},
    {"Stierenkaart": "kg protein",                            "Titel in bestand": "Official Production Evaluation in this Country KG eiwit",                                   "Formule": None},
    {"Stierenkaart": "%reliability conformation traits",      "Titel in bestand": "%betrouwbaarheid (exterieur-index)",                                                       "Formule": None},
    {"Stierenkaart": "frame",                                 "Titel in bestand": "GENERAL CHARACTERISTICS frame",                                                            "Formule": "/100"},
    {"Stierenkaart": "udder",                                 "Titel in bestand": "GENERAL CHARACTERISTICS uier",                                                             "Formule": "/100"},
    {"Stierenkaart": "feet & legs",                           "Titel in bestand": "GENERAL CHARACTERISTICS benen",                                                           "Formule": "/100"},
    {"Stierenkaart": "final score",                           "Titel in bestand": "GENERAL CHARACTERISTICS totaal (Exterieur-index)",                                         "Formule": "/100"},
    {"Stierenkaart": "stature",                               "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY hoogtemaat",                               "Formule": "/100"},
    {"Stierenkaart": "chestwidth",                            "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorhand",                              "Formule": "/100"},
    {"Stierenkaart": "body depth",                            "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY inhoud",                                "Formule": "/100"},
    {"Stierenkaart": "anguliarty",                            "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY openheid",                              "Formule": "/100"},
    {"Stierenkaart": "condition score",                       "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY conditiescore",                        "Formule": "/100"},
    {"Stierenkaart": "rump angle",                            "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY kruisligging",                         "Formule": "/100"},
    {"Stierenkaart": "rump width",                            "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY kruisbreedte",                         "Formule": "/100"},
    {"Stierenkaart": "rear legs rear view",                   "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achteraanzicht benen",             "Formule": "/100"},
    {"Stierenkaart": "rear leg set",                          "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY beenstand zij",                        "Formule": "/100"},
    {"Stierenkaart": "foot angle",                            "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY klauwhoek",                           "Formule": "/100"},
    {"Stierenkaart": "front feet orientation",                "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorbeenstand",                      "Formule": "/100"},
    {"Stierenkaart": "locomotion",                            "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY beengebruik",                          "Formule": "/100"},
    {"Stierenkaart": "fore udder attachment",                 "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY vooruieraanhechting",                "Formule": "/100"},
    {"Stierenkaart": "fore teat placement",                   "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorspeenplaatsing",                "Formule": "/100"},
    {"Stierenkaart": "teat length",                           "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY speenlengte",                        "Formule": "/100"},
    {"Stierenkaart": "udder depth",                           "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY uierdiepte",                          "Formule": "/100"},
    {"Stierenkaart": "rear udder height",                     "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achteruierhoogte",                   "Formule": "/100"},
    {"Stierenkaart": "central ligament",                      "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY ophangband",                         "Formule": "/100"},
    {"Stierenkaart": "rear teat placement",                   "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achterspeenplaatsing",               "Formule": "/100"},
    {"Stierenkaart": "persistency",                           "Titel in bestand": "",                                                                                      "Formule": None},
    {"Stierenkaart": "calving ease",                          "Titel in bestand": "OFFICIAL CALVING EASE EVALUATION IN THIS COUNTRY geboortegemak",                "Formule": "/100"},
    {"Stierenkaart": "milking speed",                         "Titel in bestand": "OFFICIAL MILKING SPEED AND TEMPERAMENT EVALUATION IN THIS COUNTRY melksnelheid","Formule": "/100"},
    {"Stierenkaart": "somatic cell count",                    "Titel in bestand": "OFFICIAL SOMATIC CELL COUNT EVALUATION IN THIS COUNTRY celgetal",              "Formule": "/100"},
    {"Stierenkaart": "female fertility",                      "Titel in bestand": "OFFICIAL FEMALE FERTILITY EVALUATION IN THIS COUNTRY vruchtbaarheid",             "Formule": "/100"},
    {"Stierenkaart": "temperament",                           "Titel in bestand": "OFFICIAL MILKING SPEED AND TEMPERAMENT EVALUATION IN THIS COUNTRY karakter",  "Formule": "/100"},
    {"Stierenkaart": "maturity rate",                         "Titel in bestand": "OFFICIAL CALVING EASE EVALUATION IN THIS COUNTRY laatrijpheid",             "Formule": "/100"},
    {"Stierenkaart": "hoofhealth",                            "Titel in bestand": "OFFICIAL CLAW HEALTH EVALUATION IN THIS COUNTRY klauwgezondheid",         "Formule": "/100"},
    {"Stierenkaart": "Beta caseine",                          "Titel in bestand": "Betacasine",                                                                    "Formule": None},
    {"Stierenkaart": "Kappa caseine",                         "Titel in bestand": "Kappa-caseine",                                                                 "Formule": None},
    {"Stierenkaart": "Superbevruchter",                       "Titel in bestand": "Superbevruchter",                                                               "Formule": None},
]

def load_excel(file):
    try:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Fout bij laden Excel: {e}")
        return None

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

def create_top5_table(df):
    fokwaarden = ["geboortegemak","celgetal","vruchtbaarheid","klauwgezondheid","uier","benen"]
    blocks = []
    if df.empty:
        return pd.DataFrame()
    df["Ras_clean"] = df["Ras"].astype(str).str.strip().str.lower()
    df = df[df["Ras_clean"].isin(["holstein zwartbont","holstein zwartbont + rf","red holstein"])].copy()
    for fok in fokwaarden:
        if fok not in df.columns:
            df[fok] = pd.NA
        header = {
            "Fokwaarde": fok,
            "zwartbont_stier":"Stier","zwartbont_value":"Waarde",
            "roodbont_stier":"Stier","roodbont_value":"Waarde"
        }
        blocks.append(header)
        df_z = df[df["Ras_clean"].str.contains("zwartbont")]
        df_r = df[df["Ras_clean"].str.contains("red holstein")]
        df_z[fok] = pd.to_numeric(df_z[fok], errors="coerce")
        df_r[fok] = pd.to_numeric(df_r[fok], errors="coerce")
        df_z = df_z.sort_values(fok,ascending=False).head(5)
        df_r = df_r.sort_values(fok,ascending=False).head(5)
        for i in range(5):
            row = {"Fokwaarde":"","zwartbont_stier":"","zwartbont_value":"","roodbont_stier":"","roodbont_value":""}
            if i < len(df_z):
                row["zwartbont_stier"] = df_z.iloc[i]["naam"]
                row["zwartbont_value"] = df_z.iloc[i][fok]
            if i < len(df_r):
                row["roodbont_stier"] = df_r.iloc[i]["naam"]
                row["roodbont_value"] = df_r.iloc[i][fok]
            blocks.append(row)
        blocks.append({"Fokwaarde":"","zwartbont_stier":"","zwartbont_value":"","roodbont_stier":"","roodbont_value":""})
    return pd.DataFrame(blocks)

def main():
    st.set_page_config(layout="wide")
    st.title("Stierenkaart Generator (PIM versie)")

    # Regio-keuze
    regio = st.radio("Kies regio", ["Nederland","Canada"])
    mapping_table = mapping_table_pim if regio=="Nederland" else mapping_table_ca

    # Bestand uploaden
    uploaded = st.file_uploader("Upload PIM K.I. Samen.xlsx", type="xlsx")
    if not uploaded:
        st.info("Upload eerst het PIM-bestand.")
        return

    df_raw = load_excel(uploaded)
    if df_raw is None:
        return

    # Mapping met foutmelding per missende key
    final_data = {}
    for i, m in enumerate(mapping_table):
        try:
            std          = m["Stierenkaart"]
            kolom_titel  = m["Titel in bestand"]
            formule      = m["Formule"]
        except KeyError as e:
            st.error(f"Mapping entry {i} mist key {e}: {m}")
            continue

        if kolom_titel and kolom_titel in df_raw.columns:
            col = df_raw[kolom_titel].replace([99999,"+999"], pd.NA)
            if formule:
                col = pd.to_numeric(col, errors="coerce")
                if formule=="/10":
                    col = col/10
                elif formule=="/100":
                    col = col/100
            final_data[std] = col
        else:
            final_data[std] = ""

    df_mapped = pd.DataFrame(final_data)

    # Alleen voor Nederland: pinkenstier markeren
    if regio=="Nederland" and "geboortegemak" in df_mapped:
        df_mapped["pinkenstier"] = df_mapped["geboortegemak"].apply(lambda x: "p" if pd.notna(x) and x>100 else "")

    # Zet ki-code en Name om naar strings vóór .str-accessors
    df_mapped["ki-code"] = df_mapped["ki-code"].astype(str).str.strip().str.upper()
    if regio=="Canada" and "Name" in df_mapped:
        df_mapped["Name"] = df_mapped["Name"].astype(str)

    # Display-kolom aanmaken
    if regio=="Nederland" and "naam" in df_mapped:
        df_mapped["Display"] = df_mapped["ki-code"] + " - " + df_mapped["naam"].astype(str)
    elif regio=="Canada":
        if "Name" not in df_mapped:
            st.error("Voor Canada ontbreekt kolom 'Name' na mappen.")
            return
        df_mapped["Display"] = df_mapped["ki-code"] + " - " + df_mapped["Name"]

    # Kolomvolgorde en bestandsnaam
    if regio=="Nederland":
        volgorde = ["superbevruchter","ki-code","naam","pinkenstier"] + [
            m["Stierenkaart"] for m in mapping_table_pim
            if m["Stierenkaart"] not in ("superbevruchter","ki-code","naam")
        ]
        fname = "stierenkaart_nederland.xlsx"
    else:
        volgorde = [m["Stierenkaart"] for m in mapping_table_ca]
        fname    = "stierenkaart_canada.xlsx"

    bestaand = [c for c in volgorde if c in df_mapped.columns]
    rest     = [c for c in df_mapped.columns if c not in bestaand]
    df_mapped = df_mapped[bestaand + rest]

    # Selectie en tonen
    keuze = st.multiselect("Selecteer stieren:", df_mapped["Display"])
    if not keuze:
        st.info("Selecteer minimaal één stier om te tonen en te downloaden.")
        return

    codes = [k.split(" - ")[0] for k in keuze]
    df_sel = df_mapped[df_mapped["ki-code"].isin(codes)].copy()
    df_sel = custom_sort_ras(df_sel)
    st.subheader("Geselecteerde stieren")
    st.dataframe(df_sel, use_container_width=True)

    df_top5 = create_top5_table(df_sel)
    if not df_top5.empty:
        st.subheader("Top 5-tabellen per fokwaarde")
        st.dataframe(df_top5, use_container_width=True)

    # Excel schrijven en downloaden
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_sel.to_excel(writer, sheet_name="Stierenkaart", index=False)
        if not df_top5.empty:
            df_top5.to_excel(writer, sheet_name="Top5_per_ras", index=False)

    st.download_button(
        label="Download selectie + Top 5-tabellen",
        data=buf.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    main()
