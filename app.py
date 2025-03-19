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

# Bestandsuploaders
uploaded_crv = st.file_uploader("Upload Bronbestand CRV DEC2024.xlsx", type=["xlsx"], key="crv")
uploaded_pim = st.file_uploader("Upload PIM K.I. Samen.xlsx", type=["xlsx"], key="pim")
uploaded_prijslijst = st.file_uploader("Upload Prijslijst.xlsx", type=["xlsx"], key="prijslijst")
uploaded_joop = st.file_uploader("Upload Bronbestand Joop Olieman.xlsx", type=["xlsx"], key="joop")

debug_mode = st.checkbox("Activeer debug", value=False)

if "df_stierenkaart" not in st.session_state:
    st.session_state.df_stierenkaart = None

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
            # Normaliseer de KI-codes
            df_crv["KI_Code"] = df_crv["KI-Code"].astype(str).str.upper().str.strip()
            df_pim["KI_Code"] = df_pim["Stiercode NL / KI code"].astype(str).str.upper().str.strip()
            df_prijslijst["KI_Code"] = df_prijslijst["Artikelnr."].astype(str).str.upper().str.strip()
            df_joop["KI_Code"] = df_joop["Kicode"].astype(str).str.upper().str.strip()
            
            # In het PIM-bestand: hernoem de kolom "PFW code" (case-insensitief) naar "PFW_pim"
            pfw_col = None
            for col in df_pim.columns:
                if col.lower() == "pfw code":
                    pfw_col = col
                    break
            if pfw_col:
                df_pim.rename(columns={pfw_col: "PFW_pim"}, inplace=True)
                df_pim["PFW_pim"] = df_pim["PFW_pim"].astype(str).str.strip()
            else:
                st.warning("Kolom 'PFW code' niet gevonden in het pimbestand.")
            
            # In het Joop-bestand: zoek naar de TIP-kolom
            tip_col = None
            for col in df_joop.columns:
                if col.strip().upper() == "TIP":
                    tip_col = col
                    break
            if not tip_col:
                for col in df_joop.columns:
                    if col.strip().upper().startswith("TIP"):
                        tip_col = col
                        break
            if tip_col:
                if tip_col != "TIP":
                    df_joop.rename(columns={tip_col: "TIP"}, inplace=True)
                df_joop["TIP"] = df_joop["TIP"].astype(str).str.strip()
            else:
                st.warning("Kolom 'TIP' niet gevonden in het Joop-bestand.")
            
            # Voeg een tijdelijke key toe (op basis van KI_Code) voor de merge
            for df_temp in [df_crv, df_pim, df_prijslijst, df_joop]:
                df_temp["temp_key"] = df_temp["KI_Code"]
            
            # Voer de merges uit
            df_merged = pd.merge(df_crv, df_pim, on="temp_key", how="left", suffixes=("", "_pim"))
            df_merged = pd.merge(df_merged, df_prijslijst, on="temp_key", how="left", suffixes=("", "_prijslijst"))
            df_merged = pd.merge(df_merged, df_joop, on="temp_key", how="left", suffixes=("", "_joop"))
            
            # Zorg dat de KI_Code uit het CRV-bestand behouden blijft
            if "KI_Code" in df_crv.columns:
                df_merged["KI_Code"] = df_crv["KI_Code"]
            else:
                st.error("Kolom 'KI_Code' ontbreekt in CRV-bestand.")
            
            if debug_mode:
                st.write("Debug: Kolommen in df_merged:", df_merged.columns.tolist())
                st.write("Debug: Voorbeeld data PFW_pim:", df_merged[["KI_Code", "PFW_pim"]].head())
                st.write("Debug: Voorbeeld data TIP:", df_merged[["KI_Code", "TIP"]].head())
            
            # Definieer de mapping-tabel (pas indien nodig kolomnamen aan)
            mapping_table = [
                {"Titel in bestand": "KI_Code",        "Stierenkaart": "KI-code",           "Waar te vinden": ""},
                {"Titel in bestand": "Stiernummer",     "Stierenkaart": "Stiernummer",       "Waar te vinden": ""},
                {"Titel in bestand": "Stiernaam",       "Stierenkaart": "Stier",             "Waar te vinden": ""},
                {"Titel in bestand": "Erf-fact",        "Stierenkaart": "Erf-fact",          "Waar te vinden": ""},
                {"Titel in bestand": "Vader",           "Stierenkaart": "Afstamming V",      "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "M-vader",         "Stierenkaart": "Afstamming MV",     "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "PFW_pim",         "Stierenkaart": "PFW",             "Waar te vinden": "PIM K.I. SAMEN"},
                {"Titel in bestand": "AAa code",        "Stierenkaart": "aAa",             "Waar te vinden": "PIM K.I. SAMEN"},
                {"Titel in bestand": "Betacasine",      "Stierenkaart": "beta caseïne",    "Waar te vinden": "PIM K.I. SAMEN"},
                {"Titel in bestand": "Kappa-caseine",   "Stierenkaart": "kappa Caseïne",   "Waar te vinden": "PIM K.I. SAMEN"},
                {"Titel in bestand": "Prijs",           "Stierenkaart": "Prijs",           "Waar te vinden": "Prijslijst"},
                {"Titel in bestand": "Prijs gesekst",   "Stierenkaart": "Prijs gesekst",   "Waar te vinden": "Prijslijst"},
                {"Titel in bestand": "Bt_1",            "Stierenkaart": "% betrouwbaarheid", "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "kgM",             "Stierenkaart": "kg melk",         "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "%V",              "Stierenkaart": "% vet",           "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "%E",              "Stierenkaart": "% eiwit",         "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "kgV",             "Stierenkaart": "kg vet",          "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "kgE",             "Stierenkaart": "kg eiwit",        "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "INET",            "Stierenkaart": "INET",            "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "NVI",             "Stierenkaart": "NVI",             "Waar te vinden": "Bronbestand CRV"},
                {"Titel in bestand": "TIP",             "Stierenkaart": "TIP",             "Waar te vinden": "Bronbestand Joop Olieman"},
                {"Titel in bestand": "Rasomschrijving", "Stierenkaart": "Ras",             "Waar te vinden": "PIM K.I. SAMEN"}
            ]
            
            final_data = {}
            for mapping in mapping_table:
                titel = mapping["Titel in bestand"]
                std_naam = mapping["Stierenkaart"]
                final_data[std_naam] = df_merged.get(titel, "")
            
            df_stierenkaart = pd.DataFrame(final_data)
            df_stierenkaart.fillna("", inplace=True)
            
            st.session_state.df_stierenkaart = df_stierenkaart

# SELECTIE UI met groepering per Ras
if st.session_state.get("df_stierenkaart") is not None:
    df_stierenkaart = st.session_state.df_stierenkaart

    # Maak een Display-kolom in de vorm "KI-code - Stier"
    df_stierenkaart["Display"] = df_stierenkaart["KI-code"] + " - " + df_stierenkaart["Stier"]

    # Bulk-selectie (optioneel)
    bulk_file = st.file_uploader("Upload bulk selectie bestand (KI-code kolom A)", type=["xlsx"], key="bulk")
    bulk_selected_codes = []
    if bulk_file:
        df_bulk = pd.read_excel(bulk_file)
        bulk_selected_codes = df_bulk.iloc[:, 0].astype(str).str.upper().str.strip().tolist()

    # Groepeer de Display-opties per Ras
    grouped_options = {}
    for _, row in df_stierenkaart.iterrows():
        breed = row["Ras"]
        display = row["Display"]
        if breed not in grouped_options:
            grouped_options[breed] = []
        grouped_options[breed].append(display)
    # Verwijder dubbele waarden en sorteer elke groep
    for breed in grouped_options:
        grouped_options[breed] = sorted(list(set(grouped_options[breed])))

    # Definieer een custom order voor de rassen:
    order_map = {
        "Holstein zwartbont": 1,
        "Red holstein": 2
    }
    # Sorteer de rassen: eerst de rassen uit order_map, dan de overige
    sorted_breeds = sorted(grouped_options.keys(), key=lambda x: order_map.get(x, 3))
    
    st.markdown("### Selectie per ras")
    selected_displays = []
    for breed in sorted_breeds:
        st.markdown(f"**{breed}**")
        sel = st.multiselect(f"Selecteer stieren ({breed})", options=grouped_options[breed], key=f"ms_{breed}")
        selected_displays.extend(sel)
    
    # Combineer de gemaakte selectie met de bulk-selectie
    manual_selected_codes = [item.split(" - ")[0] for item in selected_displays]
    combined_codes = sorted(set(bulk_selected_codes + manual_selected_codes))
    mapping_dict = dict(zip(df_stierenkaart["KI-code"], df_stierenkaart["Display"]))
    valid_final_display = [mapping_dict.get(code) for code in combined_codes if mapping_dict.get(code) is not None]
    
    # Optioneel: een gecombineerde multiselect
    final_combined_display = st.multiselect("Gecombineerde selectie:", options=list(mapping_dict.values()), default=valid_final_display)
    final_selected_codes = [item.split(" - ")[0] for item in final_combined_display]
    
    # Filter de DataFrame op basis van de uiteindelijke KI-codes
    df_selected = df_stierenkaart[df_stierenkaart["KI-code"].isin(final_selected_codes)]
    df_overig = df_stierenkaart[~df_stierenkaart["KI-code"].isin(final_selected_codes)]
    
    # Custom sortering op Ras
    def custom_sort_ras(df):
        order_map = {"Holstein zwartbont": 1, "Red holstein": 2}
        if "Ras" not in df.columns:
            df["Ras"] = ""
        df["ras_sort"] = df["Ras"].map(order_map).fillna(3)
        df_sorted = df.sort_values(by=["ras_sort", "Stier"], ascending=True)
        df_sorted.drop(columns=["ras_sort"], inplace=True)
        return df_sorted
    
    df_selected = custom_sort_ras(df_selected)
    df_overig = custom_sort_ras(df_overig)
    
    # Exporteer naar Excel
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
