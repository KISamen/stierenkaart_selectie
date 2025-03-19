import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide")
st.title("Stierenkaart Generator")

st.markdown("""
**Upload de volgende bestanden:**
- **Bronbestand CRV DEC2024.xlsx**
- **PIM K.I. Samen.xlsx**
- **Prijslijst.xlsx**
- **Bronbestand Joop Olieman.xlsx**

**Bulk-selectie:**  
Excelbestand met KI-code in kolom A.
""")

uploaded_crv = st.file_uploader("Upload Bronbestand CRV DEC2024.xlsx", type=["xlsx"], key="crv")
uploaded_pim = st.file_uploader("Upload PIM K.I. Samen.xlsx", type=["xlsx"], key="pim")
uploaded_prijslijst = st.file_uploader("Upload Prijslijst.xlsx", type=["xlsx"], key="prijslijst")
uploaded_joop = st.file_uploader("Upload Bronbestand Joop Olieman.xlsx", type=["xlsx"], key="joop")

debug_mode = st.checkbox("Activeer debug", value=False)

if "df_stierenkaart" not in st.session_state:
    st.session_state.df_stierenkaart = None
if "df_mapping" not in st.session_state:
    st.session_state.df_mapping = None
if "bulk_selected" not in st.session_state:
    st.session_state.bulk_selected = []

def load_excel(file):
    try:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Fout bij laden bestand: {e}")
        return None

if st.button("Genereer Stierenkaart"):
    if not (uploaded_crv and uploaded_pim and uploaded_prijslijst and uploaded_joop):
        st.error("Upload alle bestanden!")
    else:
        df_crv = load_excel(uploaded_crv)
        df_pim = load_excel(uploaded_pim)
        df_prijslijst = load_excel(uploaded_prijslijst)
        df_joop = load_excel(uploaded_joop)

        if any(df is None for df in [df_crv, df_pim, df_prijslijst, df_joop]):
            st.error("Fout bij laden bestanden.")
        else:
            # Normaliseer KI-codes
            df_crv["KI_Code"] = df_crv["KI-Code"].astype(str).str.upper().str.strip()
            df_pim["KI_Code"] = df_pim["Stiercode NL / KI code"].astype(str).str.upper().str.strip()
            df_prijslijst["KI_Code"] = df_prijslijst["Artikelnr."].astype(str).str.upper().str.strip()
            df_joop["KI_Code"] = df_joop["Kicode"].astype(str).str.upper().str.strip()

            # Zoek case-onafhankelijk naar de kolom "PFW code" en hernoem deze naar "PFW_pim"
            pfw_col = None
            for col in df_pim.columns:
                if col.lower() == "pfw code":
                    pfw_col = col
                    break
            if pfw_col:
                df_pim.rename(columns={pfw_col: "PFW_pim"}, inplace=True)
                # Converteer de PFW_pim kolom expliciet naar string en strip spaties
                df_pim["PFW_pim"] = df_pim["PFW_pim"].astype(str).str.strip()
            else:
                st.warning("Kolom 'PFW code' niet gevonden in het pimbestand.")

            # Voeg een tijdelijke key toe voor de merges
            for df in [df_crv, df_pim, df_prijslijst, df_joop]:
                df["temp_key"] = df["KI_Code"]

            # Merge de dataframes
            df_merged = pd.merge(df_crv, df_pim, on="temp_key", how="left", suffixes=("", "_pim"))
            df_merged = pd.merge(df_merged, df_prijslijst, on="temp_key", how="left", suffixes=("", "_prijslijst"))
            df_merged = pd.merge(df_merged, df_joop, on="temp_key", how="left", suffixes=("", "_joop"))

            # Indien CRV de KI_Code kolom bevat, zorg ervoor dat deze wordt behouden
            if "KI_Code" in df_crv.columns:
                df_merged["KI_Code"] = df_crv["KI_Code"]
            else:
                st.error("Kolom 'KI_Code' ontbreekt in CRV-bestand.")

            if debug_mode:
                st.write("Debug: Kolommen in df_merged", df_merged.columns.tolist())
                if "PFW_pim" in df_merged.columns:
                    st.write("Debug: Voorbeeld data PFW_pim", df_merged[["KI_Code", "PFW_pim"]].head())
                else:
                    st.write("Debug: 'PFW_pim' kolom ontbreekt in df_merged.")

            # Definieer de mapping-tabel; let op de aanpassing voor de PFW kolom
            mapping_table = [
                {"Titel in bestand": "KI_Code",        "Stierenkaart": "KI-code",           "Waar te vinden": ""},
                {"Titel in bestand": "Stiernummer",      "Stierenkaart": "Stiernummer",         "Waar te vinden": ""},
                {"Titel in bestand": "Stiernaam",        "Stierenkaart": "Stier",               "Waar te vinden": ""},
                {"Titel in bestand": "Erf-fact",         "Stierenkaart": "Erf-fact",            "Waar te vinden": ""},
                {"Titel in bestand": "Vader",            "Stierenkaart": "Afstamming V",        "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "M-vader",          "Stierenkaart": "Afstamming MV",       "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "PFW_pim",          "Stierenkaart": "PFW",               "Waar te vinden": "PIM K.I. SAMEN"},
                {"Titel in bestand": "AAa code",         "Stierenkaart": "aAa",               "Waar te vinden": "PIM K.I. SAMEN"},
                {"Titel in bestand": "Betacasine",       "Stierenkaart": "beta caseïne",      "Waar te vinden": "PIM K.I. SAMEN"},
                {"Titel in bestand": "Kappa-caseine",    "Stierenkaart": "kappa Caseïne",     "Waar te vinden": "PIM K.I. SAMEN"},
                {"Titel in bestand": "Prijs",            "Stierenkaart": "Prijs",             "Waar te vinden": "Prijslijst"},
                {"Titel in bestand": "Prijs gesekst",    "Stierenkaart": "Prijs gesekst",     "Waar te vinden": "Prijslijst"},
                {"Titel in bestand": "Bt_1",             "Stierenkaart": "% betrouwbaarheid", "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "kgM",              "Stierenkaart": "kg melk",           "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "%V",               "Stierenkaart": "% vet",             "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "%E",               "Stierenkaart": "% eiwit",           "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "kgV",              "Stierenkaart": "kg vet",            "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "kgE",              "Stierenkaart": "kg eiwit",          "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "INET",             "Stierenkaart": "INET",              "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "NVI",              "Stierenkaart": "NVI",               "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "TIP",              "Stierenkaart": "TIP",               "Waar te vinden": "Bronbestand Joop Olieman"},
                {"Titel in bestand": "Bt_5",             "Stierenkaart": "% betrouwbaar",     "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "F",                "Stierenkaart": "frame",             "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "U",                "Stierenkaart": "uier",              "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "B_6",              "Stierenkaart": "benen",             "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "Ext",              "Stierenkaart": "totaal",            "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "HT",               "Stierenkaart": "hoogtemaat",        "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "VH",               "Stierenkaart": "voorhand",          "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "IH",               "Stierenkaart": "inhoud",            "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "OH",               "Stierenkaart": "openheid",          "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "CS",               "Stierenkaart": "conditie score",    "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "KL",               "Stierenkaart": "kruisligging",      "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "KB",               "Stierenkaart": "kruisbreedte",      "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "BA",               "Stierenkaart": "beenstand achter",  "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "BZ",               "Stierenkaart": "beenstand zij",     "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "KH",               "Stierenkaart": "klauwhoek",         "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "VB",               "Stierenkaart": "voorbeenstand",     "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "BG",               "Stierenkaart": "beengebruik",       "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "VA",               "Stierenkaart": "vooruieraanhechting","Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "VP",               "Stierenkaart": "voorspeenplaatsing", "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "SL",               "Stierenkaart": "speenlengte",       "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "UD",               "Stierenkaart": "uierdiepte",        "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "AH",               "Stierenkaart": "achteruierhoogte",  "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "OB",               "Stierenkaart": "ophangband",        "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "AP",               "Stierenkaart": "achterspeenplaatsing", "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "Geb",              "Stierenkaart": "Geboortegemak",      "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "MS",               "Stierenkaart": "melksnelheid",      "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "Cgt",              "Stierenkaart": "celgetal",         "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "Vru",              "Stierenkaart": "vruchtbaarheid",   "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "KA",               "Stierenkaart": "karakter",         "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "Ltrh",             "Stierenkaart": "laatrijpheid",     "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "Pers",             "Stierenkaart": "Persistentie",     "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "Kgh",              "Stierenkaart": "klauwgezondheid",  "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "Lvd",              "Stierenkaart": "levensduur",       "Waar te vinden": "Bronbestand CRV"}
            ]

            final_data = {}
            for mapping in mapping_table:
                titel = mapping["Titel in bestand"]
                std_naam = mapping["Stierenkaart"]
                final_data[std_naam] = df_merged.get(titel, "")

            df_stierenkaart = pd.DataFrame(final_data)
            df_stierenkaart.fillna("", inplace=True)

            st.session_state.df_stierenkaart = df_stierenkaart

if st.session_state.get("df_stierenkaart") is not None:
    df_stierenkaart = st.session_state.df_stierenkaart

    df_stierenkaart["Display"] = df_stierenkaart["KI-code"] + " - " + df_stierenkaart["Stier"]
    options = sorted(df_stierenkaart["Display"].dropna().unique().tolist())

    bulk_file = st.file_uploader("Upload bulk selectie bestand (KI-code kolom A)", type=["xlsx"], key="bulk")
    bulk_selected_codes = []

    if bulk_file:
        df_bulk = pd.read_excel(bulk_file)
        bulk_selected_codes = df_bulk.iloc[:, 0].astype(str).str.upper().str.strip().tolist()

    manual_selected_display = st.multiselect("Voeg extra stieren toe:", options=options)
    manual_selected_codes = [item.split(" - ")[0] for item in manual_selected_display]

    combined_codes = sorted(set(bulk_selected_codes + manual_selected_codes))
    mapping_dict = dict(zip(df_stierenkaart["KI-code"], df_stierenkaart["Display"]))
    valid_final_display = [mapping_dict.get(code) for code in combined_codes if mapping_dict.get(code) in options]

    final_combined_display = st.multiselect("Gecombineerde selectie:", options=options, default=valid_final_display)
    final_selected_codes = [item.split(" - ")[0] for item in final_combined_display]

    df_selected = df_stierenkaart[df_stierenkaart["KI-code"].isin(final_selected_codes)]
    df_overig = df_stierenkaart[~df_stierenkaart["KI-code"].isin(final_selected_codes)]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_selected.to_excel(writer, sheet_name='Stierenkaart', index=False)
        df_overig.to_excel(writer, sheet_name='Overige stieren', index=False)

    st.download_button(
        label="Download stierenkaart Excel",
        data=output.getvalue(),
        file_name="stierenkaart.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
