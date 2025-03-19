import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide")

st.title("Stierenkaart Generator")

st.markdown("""
**Upload de volgende bestanden:**
- **Bronbestand CRV DEC2024.xlsx** (bevat kolommen zoals "KI-Code", "Vader", "M-vader", etc.)
- **PIM K.I. Samen.xlsx** (bevat kolommen zoals "Stiercode NL / KI code" en PIM-data, bv. PFW, AAa code, Betacasine, Kappa-caseine)
- **Prijslijst.xlsx** (kolom: "Artikelnr.")
- **Bronbestand Joop Olieman.xlsx** (bevat onder andere de kolom "Kicode")

**Bulk-selectie:**  
Je kunt een Excelbestand uploaden waarin in kolom A de KI‑code staat (bijv. 782666).  
Deze KI‑codes worden gebruikt om de bijbehorende bullgegevens (Stier) op te zoeken.

De uiteindelijke selectie is de combinatie (unie) van de KI‑codes uit het bulkbestand en extra handmatig geselecteerde KI‑codes.  
Ontbrekende data blijft leeg in de export.
""")

# 1. Upload de 4 benodigde bestanden
uploaded_crv = st.file_uploader("Upload Bronbestand CRV DEC2024.xlsx", type=["xlsx"], key="crv")
uploaded_pim = st.file_uploader("Upload PIM K.I. Samen.xlsx", type=["xlsx"], key="pim")
uploaded_prijslijst = st.file_uploader("Upload Prijslijst.xlsx", type=["xlsx"], key="prijslijst")
uploaded_joop = st.file_uploader("Upload Bronbestand Joop Olieman.xlsx", type=["xlsx"], key="joop")

# Checkbox voor debugmodus
debug_mode = st.checkbox("Activeer debug", value=False)

# Session state om stierenkaart en mapping op te slaan
if "df_stierenkaart" not in st.session_state:
    st.session_state.df_stierenkaart = None
if "df_mapping" not in st.session_state:
    st.session_state.df_mapping = None
if "bulk_selected" not in st.session_state:
    st.session_state.bulk_selected = []

def load_excel(file):
    """Lees een Excel-bestand in en trim de kolomnamen."""
    try:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Fout bij het laden van het bestand: {e}")
        return None

# 2. Verwerk de bestanden en bouw de stierenkaart
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
            # Debug: kolomnamen in PIM
            if debug_mode:
                st.write("Kolommen in PIM bestand:", df_pim.columns.tolist())
            
            # Normaliseer de KI-code in elk bestand
            df_crv["KI_Code"] = df_crv["KI-Code"].astype(str).str.upper().str.strip()
            df_pim["KI_Code"] = df_pim["Stiercode NL / KI code"].astype(str).str.upper().str.strip()
            df_prijslijst["KI_Code"] = df_prijslijst["Artikelnr."].astype(str).str.upper().str.strip()
            df_joop["KI_Code"] = df_joop["Kicode"].astype(str).str.upper().str.strip()
            
            if debug_mode:
                st.write("Voorbeeld KI_Codes in CRV:", df_crv["KI_Code"].head().tolist())
                st.write("Voorbeeld KI_Codes in PIM:", df_pim["KI_Code"].head().tolist())
            
            # Voeg een temp_key toe voor de merges
            for df in [df_crv, df_pim, df_prijslijst, df_joop]:
                df["temp_key"] = df["KI_Code"]
            
            common_keys = set(df_crv["temp_key"]).intersection(set(df_pim["temp_key"]))
            if debug_mode:
                st.write("Aantal gemeenschappelijke KI-codes tussen CRV en PIM:", len(common_keys))
            
            # Merge de dataframes (CRV als basis)
            df_merged = pd.merge(df_crv, df_pim, on="temp_key", how="left", suffixes=("", "_pim"))
            df_merged = pd.merge(df_merged, df_prijslijst, on="temp_key", how="left", suffixes=("", "_prijslijst"))
            df_merged = pd.merge(df_merged, df_joop, on="temp_key", how="left", suffixes=("", "_joop"))
            df_merged["KI_Code"] = df_crv["KI_Code"]
            df_merged.drop(columns=["temp_key"], inplace=True)
            
            if debug_mode:
                st.write("Kolommen in merged dataframe:", df_merged.columns.tolist())
            
            # Mappingtabel: we gebruiken "KI_Code" als bron, en willen "KI-code" in de output
            mapping_table = [
                {"Titel in bestand": "KI_Code",        "Stierenkaart": "KI-code",           "Waar te vinden": ""},
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
            
            # Bouw de uiteindelijke DataFrame op basis van de mappingtabel
            final_data = {}
            for mapping in mapping_table:
                titel = mapping["Titel in bestand"]
                std_naam = mapping["Stierenkaart"]
                bron = mapping["Waar te vinden"]
                
                # Als "Stiernaam" ontbreekt maar "Stier" wel bestaat
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
            
            # Optioneel: Als je lege velden liever echt leeg in Excel wilt in plaats van NaN, doe:
            # df_stierenkaart.fillna("", inplace=True)
            
            df_mapping = pd.DataFrame(mapping_table)
            
            # Normaliseer de bullnamen (Stier) naar hoofdletters
            if "Stier" in df_stierenkaart.columns:
                df_stierenkaart["Stier"] = df_stierenkaart["Stier"].astype(str).str.strip().str.upper()
            
            # Sla op in session state
            st.session_state.df_stierenkaart = df_stierenkaart
            st.session_state.df_mapping = df_mapping

# 3. Als de stierenkaart beschikbaar is, toon de selectie-interface
if st.session_state.get("df_stierenkaart") is not None:
    df_stierenkaart = st.session_state.df_stierenkaart
    df_mapping = st.session_state.df_mapping

    # Maak een "Display" kolom voor handmatige selectie: "KI-code - Stier"
    code_col = "KI-code" if "KI-code" in df_stierenkaart.columns else "KI_Code"
    df_stierenkaart["Display"] = df_stierenkaart[code_col].astype(str) + " - " + df_stierenkaart["Stier"].astype(str)
    
    # Alle mogelijke opties voor handmatige selectie
    options = sorted(df_stierenkaart["Display"].dropna().unique().tolist())
    
    # 3.1 Bulk-upload (met KI-codes in kolom A)
    bulk_file = st.file_uploader("Upload bulk selectie bestand (met KI-code in kolom A)", type=["xlsx"], key="bulk")
    bulk_selected_codes = []
    if bulk_file is not None:
        try:
            df_bulk = pd.read_excel(bulk_file)
            df_bulk.columns = df_bulk.columns.str.strip()
            # Neem de eerste kolom als KI-code
            bulk_codes = df_bulk.iloc[:, 0].dropna().astype(str).str.upper().str.strip().unique().tolist()
            bulk_selected_codes = sorted(bulk_codes)
            st.session_state.bulk_selected = bulk_selected_codes
            
            if debug_mode:
                st.write("Bulk selectie (KI-codes uit bestand):", bulk_selected_codes)
        except Exception as e:
            st.error("Fout bij het laden van het bulk selectie bestand: " + str(e))
    else:
        bulk_selected_codes = st.session_state.get("bulk_selected", [])
    
    # 3.2 Handmatige selectie
    manual_selected_display = st.multiselect(
        "Voeg extra stieren toe (handmatige selectie):",
        options=options,
        default=[]
    )
    manual_selected_codes = [item.split(" - ")[0] for item in manual_selected_display]
    
    # 3.3 Combineer bulk- en handmatige selectie
    final_selected_codes = sorted(set(bulk_selected_codes).union(set(manual_selected_codes)))
    
    # Maak van die codes de "Display" strings
    mapping_dict = dict(zip(df_stierenkaart[code_col], df_stierenkaart["Display"]))
    final_display = [mapping_dict.get(code, code) for code in final_selected_codes]
    
    # Filter final_display op items die in options zitten, zodat er geen fout in multiselect komt
    valid_final_display = [item for item in final_display if item in options]
    
    # 3.4 Extra multiselect-widget om de gecombineerde lijst te bewerken
    final_combined_display = st.multiselect(
        "Bewerk de gecombineerde stierenlijst:",
        options=options,
        default=valid_final_display
    )
    # Haal de definitieve KI-codes uit de bewerkte lijst
    final_selected_codes = [item.split(" - ")[0] for item in final_combined_display]
    
    if debug_mode:
        st.write("Gecombineerde selectie (KI-code - Stier):", final_combined_display)
    
    # 4. Filter de DataFrame op basis van final_selected_codes
    df_selected = df_stierenkaart[df_stierenkaart[code_col].isin(final_selected_codes)]
    df_overig = df_stierenkaart[~df_stierenkaart[code_col].isin(final_selected_codes)]
    
    if debug_mode:
        st.write("Aantal geselecteerde rijen:", len(df_selected))
        st.write("Aantal overige rijen:", len(df_overig))
    
    # 5. Exporteer naar Excel (Stierenkaart, Overige stieren, Mapping)
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
