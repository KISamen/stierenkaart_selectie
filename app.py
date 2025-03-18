import streamlit as st
import pandas as pd
import io

st.title("Stierenkaart Generator")

st.markdown("""
Upload de volgende bestanden:
- **Bronbestand CRV DEC2024.xlsx** (bevat kolommen zoals "KI-Code", "Vader", "M-vader", etc.)
- **PIM K.I. Samen.xlsx** (bevat kolommen zoals "Stiercode NL / KI code" en PIM-data: PFW, AAa code, Betacasine, Kappa-caseine)
- **Prijslijst.xlsx** (kolom: "Artikelnr.")
- **Bronbestand Joop Olieman.xlsx** (bevat onder andere de kolom "Kicode")

Na verwerking kun je stieren selecteren voor de hoofd-export.
Je kunt een bulk-selectiebestand uploaden met stiergegevens.  
Belangrijk: in dit bulkbestand moet de stier worden gematcht op basis van de KI-code (kolom A) – niet op de naam.
Als de bulkfile geen KI-code bevat, wordt er gekeken naar "Stier", "Stiernaam" of "Naam".
Ontbrekende data blijft leeg in de export.
""")

# Bewaar de uiteindelijke stierenkaart, mappingtabel en bulk selectie in de session state
if "df_stierenkaart" not in st.session_state:
    st.session_state.df_stierenkaart = None
if "df_mapping" not in st.session_state:
    st.session_state.df_mapping = None
if "bulk_selected" not in st.session_state:
    st.session_state.bulk_selected = []

# Upload de vier hoofdbestanden
uploaded_crv = st.file_uploader("Upload Bronbestand CRV DEC2024.xlsx", type=["xlsx"], key="crv")
uploaded_pim = st.file_uploader("Upload PIM K.I. Samen.xlsx", type=["xlsx"], key="pim")
uploaded_prijslijst = st.file_uploader("Upload Prijslijst.xlsx", type=["xlsx"], key="prijslijst")
uploaded_joop = st.file_uploader("Upload Bronbestand Joop Olieman.xlsx", type=["xlsx"], key="joop")

def load_excel(file):
    try:
        df = pd.read_excel(file)
        # Zorg dat kolomnamen geen overbodige spaties bevatten
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Fout bij het laden van het bestand: {e}")
        return None

# Verwerk de bestanden en bouw de stierenkaart
if st.button("Genereer Stierenkaart"):
    if not (uploaded_crv and uploaded_pim and uploaded_prijslijst and uploaded_joop):
        st.error("Zorg dat je alle vereiste bestanden uploadt!")
    else:
        df_crv = load_excel(uploaded_crv)
        df_pim = load_excel(uploaded_pim)
        df_prijslijst = load_excel(uploaded_prijslijst)
        df_joop = load_excel(uploaded_joop)
        
        if any(df is None for df in [df_crv, df_pim, df_prijslijst, df_joop]):
            st.error("Er is een fout opgetreden bij het laden van een of meerdere bestanden.")
        else:
            st.write("Kolommen in PIM bestand:", df_pim.columns.tolist())
            
            # Normaliseer de KI-code in elk bestand
            df_crv["KI_Code"] = df_crv["KI-Code"].astype(str).str.upper().str.strip()
            df_pim["KI_Code"] = df_pim["Stiercode NL / KI code"].astype(str).str.upper().str.strip()
            df_prijslijst["KI_Code"] = df_prijslijst["Artikelnr."].astype(str).str.upper().str.strip()
            df_joop["KI_Code"] = df_joop["Kicode"].astype(str).str.upper().str.strip()
            
            st.write("Voorbeeld KI_Codes in CRV:", df_crv["KI_Code"].head().tolist())
            st.write("Voorbeeld KI_Codes in PIM:", df_pim["KI_Code"].head().tolist())
            
            # Maak een tijdelijke merge-sleutel in alle bestanden
            df_crv["temp_key"] = df_crv["KI_Code"]
            df_pim["temp_key"] = df_pim["KI_Code"]
            df_prijslijst["temp_key"] = df_prijslijst["KI_Code"]
            df_joop["temp_key"] = df_joop["KI_Code"]
            
            common_keys = set(df_crv["temp_key"]).intersection(set(df_pim["temp_key"]))
            st.write("Aantal gemeenschappelijke KI-codes tussen CRV en PIM:", len(common_keys))
            
            # Merge de dataframes: gebruik CRV als basis (left join)
            df_merged = pd.merge(df_crv, df_pim, on="temp_key", how="left", suffixes=("", "_pim"))
            df_merged = pd.merge(df_merged, df_prijslijst, on="temp_key", how="left", suffixes=("", "_prijslijst"))
            df_merged = pd.merge(df_merged, df_joop, on="temp_key", how="left", suffixes=("", "_joop"))
            df_merged["KI_Code"] = df_crv["KI_Code"]
            df_merged.drop(columns=["temp_key"], inplace=True)
            
            st.write("Kolommen in merged dataframe:", df_merged.columns.tolist())
            
            # Definieer de mappingtabel
            mapping_table = [
                {"Titel in bestand": "KI-Code",        "Stierenkaart": "KI-code",           "Waar te vinden": ""},
                {"Titel in bestand": "Eigenaarscode",    "Stierenkaart": "Eigenaarscode",       "Waar te vinden": ""},
                {"Titel in bestand": "Stiernummer",      "Stierenkaart": "Stiernummer",         "Waar te vinden": ""},
                # Gebruik "Stiernaam" als eerste keuze, met fallback op "Stier"
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
            
            # Bouw de uiteindelijke DataFrame op basis van de mappingtabel
            final_data = {}
            for mapping in mapping_table:
                titel = mapping["Titel in bestand"]
                std_naam = mapping["Stierenkaart"]
                bron = mapping["Waar te vinden"]
                
                # Fallback: als we de bullnaam willen en "Stiernaam" ontbreekt maar "Stier" wel bestaat, dan gebruiken we "Stier"
                if std_naam == "Stier" and titel not in df_merged.columns and "Stier" in df_merged.columns:
                    final_data[std_naam] = df_merged["Stier"]
                    continue

                if bron == "PIM K.I. SAMEN":
                    if titel in df_merged.columns and not df_merged[titel].isnull().all():
                        final_data[std_naam] = df_merged[titel]
                    elif (titel + "_pim") in df_merged.columns and not df_merged[titel + "_pim"].isnull().all():
                        final_data[std_naam] = df_merged[titel + "_pim"]
                    else:
                        if titel in df_pim.columns:
                            df_pim_temp = df_pim.copy()
                            df_pim_temp.set_index("KI_Code", inplace=True)
                            final_data[std_naam] = df_merged["KI_Code"].map(df_pim_temp[titel])
                        else:
                            final_data[std_naam] = None
                else:
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
            
            # Normaliseer de bullnamen in de output zodat ze consistent zijn
            if "Stier" in df_stierenkaart.columns:
                df_stierenkaart["Stier"] = df_stierenkaart["Stier"].astype(str).str.strip().str.upper()
            
            st.session_state.df_stierenkaart = df_stierenkaart
            st.session_state.df_mapping = df_mapping

# Indien er een stierenkaart beschikbaar is, toon de selectie-opties
if st.session_state.get("df_stierenkaart") is not None:
    df_stierenkaart = st.session_state.df_stierenkaart
    df_mapping = st.session_state.df_mapping

    # Bepaal de beschikbare stiernamen uit de kolom "Stier"
    if "Stier" in df_stierenkaart.columns:
        options = sorted(df_stierenkaart["Stier"].dropna().unique().tolist())
    else:
        options = []
    st.write("Beschikbare stieren:", options)
    
    # Extra uploadoptie voor bulk-selectiebestand
    bulk_file = st.file_uploader("Upload bulk selectie bestand met KI-code (kolom A) en/of Stiernaam", type=["xlsx"], key="bulk")
    bulk_selected = []
    if bulk_file is not None:
        try:
            df_bulk = pd.read_excel(bulk_file)
            df_bulk.columns = df_bulk.columns.str.strip()
            # Eerst: als er een kolom "KI-code" of "KI_Code" aanwezig is, gebruiken we deze
            if "KI-code" in df_bulk.columns or "KI_Code" in df_bulk.columns:
                if "KI-code" in df_bulk.columns:
                    bulk_codes = sorted(df_bulk["KI-code"].dropna().astype(str).str.upper().str.strip().unique().tolist())
                else:
                    bulk_codes = sorted(df_bulk["KI_Code"].dropna().astype(str).str.upper().str.strip().unique().tolist())
                # Gebruik de KI-code uit de gegenereerde stierenkaart (we gaan ervan uit dat de outputkolom "KI-code" bestaat)
                # En haal de bijbehorende stiernamen op uit de kolom "Stier"
                bulk_selected = sorted(df_stierenkaart.loc[df_stierenkaart["KI-code"].isin(bulk_codes), "Stier"].unique().tolist())
            # Als er geen KI-code kolom is, probeer dan de kolom "Stier" of "Stiernaam" of "Naam"
            elif "Stier" in df_bulk.columns:
                bulk_selected = sorted(df_bulk["Stier"].dropna().astype(str).str.upper().str.strip().unique().tolist())
            elif "Stiernaam" in df_bulk.columns:
                bulk_selected = sorted(df_bulk["Stiernaam"].dropna().astype(str).str.upper().str.strip().unique().tolist())
            elif "Naam" in df_bulk.columns:
                bulk_selected = sorted(df_bulk["Naam"].dropna().astype(str).str.upper().str.strip().unique().tolist())
            st.session_state.bulk_selected = bulk_selected
            st.write("Bulk selectie (uit bestand):", bulk_selected)
        except Exception as e:
            st.error("Fout bij het laden van het bulk selectie bestand: " + str(e))
    else:
        if "bulk_selected" in st.session_state:
            bulk_selected = st.session_state.bulk_selected

    # Handmatige selectie (zonder vooraf ingevulde defaults)
    manual_selected = st.multiselect("Voeg extra stieren toe (handmatige selectie):", options=options, default=[])
    st.write("Handmatig geselecteerde stieren:", manual_selected)
    
    # Combineer de bulk en handmatige selectie (unie)
    final_selected = sorted(set(bulk_selected).union(set(manual_selected)))
    st.write("Gecombineerde selectie:", final_selected)
    
    # Filter de DataFrame op basis van de gecombineerde selectie
    df_selected = df_stierenkaart[df_stierenkaart["Stier"].isin(final_selected)]
    df_overig = df_stierenkaart[~df_stierenkaart["Stier"].isin(final_selected)]
    
    # Exporteer naar Excel met drie sheets: "Stierenkaart", "Overige stieren" en "Mapping"
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_selected.to_excel(writer, sheet_name='Stierenkaart', index=False)
        df_overig.to_excel(writer, sheet_name='Overige stieren', index=False)
        df_mapping.to_excel(writer, sheet_name='Mapping', index=False)
    excel_data = output.getvalue()
    
    st.download_button(
        label="Download gegenereerde stierenkaart",
        data=excel_data,
        file_name="stierenkaart.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.success("Excel-bestand is succesvol gegenereerd!")
