import streamlit as st
import pandas as pd
import io

st.title("Stierenkaart Generator")

st.markdown("""
Upload de volgende bestanden:
- **Bronbestand CRV DEC2024.xlsx** (bevat kolommen zoals "KI-Code", "Vader", "M-vader", etc.)
- **PIM K.I. Samen.xlsx** (bevat kolommen zoals "Stiercode NL / KI code" en PIM-data: PFW, AAa code, Betacasine, Kappa-caseine)
- **Prijslijst.xlsx** (kolom: "Artikelnr.")
- **Bronbestand Joop Olieman.xlsx** (kolom: "Kicode")
""")

# Bestanden uploaden
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
            # Debug: toon kolomnamen van de PIM-export
            st.write("Kolommen in PIM bestand:", df_pim.columns.tolist())
            
            # Normaliseer de KI-code in elk bestand (zet in hoofdletters en verwijder spaties)
            df_crv["KI_Code"] = df_crv["KI-Code"].astype(str).str.upper().str.strip()
            df_pim["KI_Code"] = df_pim["Stiercode NL / KI code"].astype(str).str.upper().str.strip()
            df_prijslijst["KI_Code"] = df_prijslijst["Artikelnr."].astype(str).str.upper().str.strip()
            df_joop["KI_Code"] = df_joop["Kicode"].astype(str).str.upper().str.strip()
            
            # Debug: toon een voorbeeld van de KI-codes in CRV en PIM
            st.write("Voorbeeld KI_Codes in CRV:", df_crv["KI_Code"].head().tolist())
            st.write("Voorbeeld KI_Codes in PIM:", df_pim["KI_Code"].head().tolist())
            
            # Maak een tijdelijke merge-sleutel in alle bestanden
            df_crv["temp_key"] = df_crv["KI_Code"]
            df_pim["temp_key"] = df_pim["KI_Code"]
            df_prijslijst["temp_key"] = df_prijslijst["KI_Code"]
            df_joop["temp_key"] = df_joop["KI_Code"]
            
            # Controleer het aantal gemeenschappelijke KI-codes
            common_keys = set(df_crv["temp_key"]).intersection(set(df_pim["temp_key"]))
            st.write("Aantal gemeenschappelijke KI-codes tussen CRV en PIM:", len(common_keys))
            
            # Merge de dataframes: gebruik CRV als basis (left join)
            df_merged = pd.merge(df_crv, df_pim, on="temp_key", how="left", suffixes=("", "_pim"))
            df_merged = pd.merge(df_merged, df_prijslijst, on="temp_key", how="left", suffixes=("", "_prijslijst"))
            df_merged = pd.merge(df_merged, df_joop, on="temp_key", how="left", suffixes=("", "_joop"))
            # Voeg ook de KI_Code toe (gebaseerd op CRV) voor latere mapping
            df_merged["KI_Code"] = df_crv["KI_Code"]
            df_merged.drop(columns=["temp_key"], inplace=True)
            
            # Debug: bekijk de kolomnamen in de merged dataframe
            st.write("Kolommen in merged dataframe:", df_merged.columns.tolist())
            
            # Definieer de mappingtabel met de aangepaste PIM-kolomnamen
            mapping_table = [
                {"Titel in bestand": "KI-Code",        "Stierenkaart": "KI-code",           "Waar te vinden": ""},
                {"Titel in bestand": "Eigenaarscode",    "Stierenkaart": "Eigenaarscode",       "Waar te vinden": ""},
                {"Titel in bestand": "Stiernummer",      "Stierenkaart": "Stiernummer",         "Waar te vinden": ""},
                {"Titel in bestand": "Stiernaam",        "Stierenkaart": "Stier",               "Waar te vinden": ""},
                {"Titel in bestand": "Erf-fact",         "Stierenkaart": "Erf-fact",            "Waar te vinden": ""},
                {"Titel in bestand": "Vader",            "Stierenkaart": "Afstamming V",        "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "M-vader",          "Stierenkaart": "Afstamming MV",       "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "PFW",              "Stierenkaart": "PFW",               "Waar te vinden": "PIM K.I. SAMEN"},
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
                bron = mapping["Waar te vinden"]
                
                if bron == "PIM K.I. SAMEN":
                    # Eerst: probeer direct uit df_merged (zonder suffix)
                    if titel in df_merged.columns and not df_merged[titel].isnull().all():
                        final_data[std_naam] = df_merged[titel]
                    # Tweede: kijk of er een kolom met suffix "_pim" aanwezig is
                    elif (titel + "_pim") in df_merged.columns and not df_merged[titel + "_pim"].isnull().all():
                        final_data[std_naam] = df_merged[titel + "_pim"]
                    else:
                        # Als de kolom niet in df_merged staat, probeer dan vanuit df_pim (indien de kolom bestaat)
                        if titel in df_pim.columns:
                            df_pim_temp = df_pim.copy()
                            df_pim_temp.set_index("KI_Code", inplace=True)
                            final_data[std_naam] = df_merged["KI_Code"].map(df_pim_temp[titel])
                        else:
                            final_data[std_naam] = None
                else:
                    # Voor overige bronnen: eerst uit df_merged, zo niet dan rechtstreeks uit de bron
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
            
            # Schrijf beide dataframes naar één Excel-bestand met twee sheets
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
