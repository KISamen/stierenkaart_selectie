import streamlit as st
import pandas as pd
import io

st.title("Stierenkaart Generator")

st.markdown("""
Upload de volgende bestanden:
- **Bronbestand CRV DEC2024.xlsx** (bevat kolommen zoals "KI-Code", "Vader", "M-vader", etc.)
- **PIM K.I. Samen.xlsx** (bevat kolommen zoals "Stiercode NL / KI code", PIM-data: PFW, aAa, Beta caseïne, Kappa caseïne)
- **Prijslijst.xlsx** (kolom: "Artikelnr.")
- **Bronbestand Joop Olieman.xlsx** (kolom: "Kicode")
""")

uploaded_crv = st.file_uploader("Upload Bronbestand CRV DEC2024.xlsx", type=["xlsx"])
uploaded_pim = st.file_uploader("Upload PIM K.I. Samen.xlsx", type=["xlsx"])
uploaded_prijslijst = st.file_uploader("Upload Prijslijst.xlsx", type=["xlsx"])
uploaded_joop = st.file_uploader("Upload Bronbestand Joop Olieman.xlsx", type=["xlsx"])

def load_excel(file):
    try:
        return pd.read_excel(file)
    except Exception as e:
        st.error(f"Fout bij het laden van het bestand: {e}")
        return None

if st.button("Genereer Stierenkaart"):
    if not (uploaded_crv and uploaded_pim and uploaded_prijslijst and uploaded_joop):
        st.error("Zorg dat je alle vereiste bestanden uploadt!")
    else:
        # Lees de bestanden in
        df_crv = load_excel(uploaded_crv)
        df_pim = load_excel(uploaded_pim)
        df_prijslijst = load_excel(uploaded_prijslijst)
        df_joop = load_excel(uploaded_joop)
        
        if any(df is None for df in [df_crv, df_pim, df_prijslijst, df_joop]):
            st.error("Er is een fout opgetreden bij het laden van een of meerdere bestanden.")
        else:
            # Voeg een kolom 'KI_Code' toe aan het CRV‑bestand (gebaseerd op "KI-Code") en maak een temp_key in alle bestanden
            df_crv["KI_Code"] = df_crv["KI-Code"].astype(str).str.strip()
            df_crv["temp_key"] = df_crv["KI_Code"]
            df_pim["temp_key"] = df_pim["Stiercode NL / KI code"].astype(str).str.strip()
            df_prijslijst["temp_key"] = df_prijslijst["Artikelnr."].astype(str).str.strip()
            df_joop["temp_key"] = df_joop["Kicode"].astype(str).str.strip()
            
            # Debug: toon aantal gemeenschappelijke KI-codes tussen CRV en PIM
            common_keys = set(df_crv["temp_key"]).intersection(set(df_pim["temp_key"]))
            st.write("Aantal gemeenschappelijke KI-codes tussen CRV en PIM:", len(common_keys))
            
            # Merge de dataframes: gebruik CRV als basis (left join)
            df_merged = pd.merge(df_crv, df_pim, on="temp_key", how="left", suffixes=("", "_pim"))
            df_merged = pd.merge(df_merged, df_prijslijst, on="temp_key", how="left", suffixes=("", "_prijslijst"))
            df_merged = pd.merge(df_merged, df_joop, on="temp_key", how="left", suffixes=("", "_joop"))
            # De temp_key is niet meer nodig
            df_merged.drop(columns=["temp_key"], inplace=True)
            
            # Debug: bekijk de kolomnamen in de merged dataframe
            st.write("Kolommen in merged dataframe:", df_merged.columns.tolist())
            
            # Mappingtabel: hier geef je per rij aan welke originele kolom (Titel in bestand)
            # moet komen te staan als welke kolom (Stierenkaart) en uit welke bron de data komt.
            mapping_table = [
                {"Titel in bestand": "KI-Code",        "Stierenkaart": "KI-code",           "Waar te vinden": ""},
                {"Titel in bestand": "Eigenaarscode",    "Stierenkaart": "Eigenaarscode",       "Waar te vinden": ""},
                {"Titel in bestand": "Stiernummer",      "Stierenkaart": "Stiernummer",         "Waar te vinden": ""},
                {"Titel in bestand": "Stiernaam",        "Stierenkaart": "Stier",               "Waar te vinden": ""},
                {"Titel in bestand": "Erf-fact",         "Stierenkaart": "Erf-fact",            "Waar te vinden": ""},
                {"Titel in bestand": "Vader",            "Stierenkaart": "Afstamming V",        "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "M-vader",          "Stierenkaart": "Afstamming MV",       "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "PFW",              "Stierenkaart": "PFW",               "Waar te vinden": "PIM K.I. SAMEN"},
                {"Titel in bestand": "aAa",              "Stierenkaart": "aAa",               "Waar te vinden": "PIM K.I. SAMEN"},
                {"Titel in bestand": "Beta caseïne",     "Stierenkaart": "beta caseïne",      "Waar te vinden": "PIM K.I. SAMEN"},
                {"Titel in bestand": "Kappa caseïne",    "Stierenkaart": "kappa Caseïne",     "Waar te vinden": "PIM K.I. SAMEN"},
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
            # Voor elke mapping halen we de data op:
            for mapping in mapping_table:
                titel = mapping["Titel in bestand"]
                std_naam = mapping["Stierenkaart"]
                bron = mapping["Waar te vinden"]
                
                if bron == "PIM K.I. SAMEN":
                    # Probeer eerst direct uit df_merged
                    if titel in df_merged.columns and not df_merged[titel].isnull().all():
                        final_data[std_naam] = df_merged[titel]
                    # Als niet, probeer kolom met suffix "_pim"
                    elif (titel + "_pim") in df_merged.columns and not df_merged[titel + "_pim"].isnull().all():
                        final_data[std_naam] = df_merged[titel + "_pim"]
                    else:
                        # Maak een mapping vanuit df_pim: indexeren op temp_key
                        df_pim_temp = df_pim.copy()
                        df_pim_temp["temp_key"] = df_pim_temp["Stiercode NL / KI code"].astype(str).str.strip()
                        df_pim_temp.set_index("temp_key", inplace=True)
                        # Gebruik de KI_Code uit df_crv (die ook in df_merged staat) om de PIM-waarde op te zoeken
                        final_data[std_naam] = df_merged["KI_Code"].map(df_pim_temp[titel])
                else:
                    # Voor overige bronnen: eerst uit df_merged, zo niet, dan rechtstreeks uit de bron
                    if titel in df_merged.columns and not df_merged[titel].isnull().all():
                        final_data[std_naam] = df_merged[titel]
                    elif bron == "Bronbestand CRV" and titel in df_crv.columns:
                        final_data[std_naam] = df_crv[titel]
                    elif bron == "Prijslijst" and titel in df_prijslijst.columns:
                        final_data[std_naam] = df_prijslijst[titel]
                    elif bron == "Bronbestand Joop Olieman" and titel in df_joop.columns:
                        final_data[std_naam] = df_joop[titel]
                    else:
                        final_data[std_naam] = None
            
            df_stierenkaart = pd.DataFrame(final_data)
            df_mapping = pd.DataFrame(mapping_table)
            
            # Schrijf het eindbestand (twee sheets: 'Stierenkaart' en 'Mapping') naar Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_stierenkaart.to_excel(writer, sheet_name='Stierenkaart', index=False)
                df_mapping.to_excel(writer, sheet_name='Mapping', index=False)
            excel_data = output.getvalue()
            
            st.download_button(
                label="Download gegenereerde stierenkaart",
                data=excel_data,
                file_name="stierenkaart.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Excel-bestand is succesvol gegenereerd!")
