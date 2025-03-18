import streamlit as st
import pandas as pd
import io

st.title("Stierenkaart Generator")
st.write("Upload de volgende bestanden om de stierenkaart te genereren:")

# Bestandsuploaders
uploaded_crv = st.file_uploader("Upload Bronbestand CRV (bijv. 'Bronbestand CRV DEC2024.xlsx')", type=["xlsx"])
uploaded_joop = st.file_uploader("Upload Bronbestand Joop Olieman (bijv. 'Bronbestand Joop Olieman.xlsx')", type=["xlsx"])
uploaded_pim = st.file_uploader("Upload Pim K.I. Samen (bijv. 'Pim K.I. Samen.xlsx')", type=["xlsx"])
uploaded_prijs = st.file_uploader("Upload Prijslijst (bijv. 'Prijslijst.xlsx')", type=["xlsx"])

if st.button("Genereer Stierenkaart"):
    if not (uploaded_crv and uploaded_joop and uploaded_pim and uploaded_prijs):
        st.warning("Zorg dat alle bestanden zijn geüpload!")
        st.stop()
    try:
        # Lees de Excel-bestanden in
        df_crv = pd.read_excel(uploaded_crv)
        df_joop = pd.read_excel(uploaded_joop)
        df_pim = pd.read_excel(uploaded_pim)
        df_prijs = pd.read_excel(uploaded_prijs)
        
        # Strip witruimtes in alle kolomnamen
        for df in [df_crv, df_joop, df_pim, df_prijs]:
            df.columns = df.columns.str.strip()
        
        # Definieer de standaard merge key voor de uiteindelijke output
        standard_key = "KI-code"
        key_variants = ["KI-code", "KI code", "KI-Code", "ki code"]
        
        # CRV-bestand: zoek naar een variant en hernoem naar standaard_key
        found_key = None
        for variant in key_variants:
            if variant in df_crv.columns:
                found_key = variant
                break
        if not found_key:
            st.error("Geen merge key gevonden in het CRV-bestand. Zorg dat een variant van 'KI-code' aanwezig is.")
            st.stop()
        if found_key != standard_key:
            df_crv.rename(columns={found_key: standard_key}, inplace=True)
        # Converteer merge key in CRV naar string, verwijder witruimtes en zet om naar hoofdletters
        df_crv[standard_key] = df_crv[standard_key].astype(str).str.strip().str.upper()
        
        st.write(f"Gebruik merge key: **{standard_key}**")
        
        # Joop Olieman-bestand: zoek naar een variant en hernoem indien gevonden
        found_key_joop = None
        for variant in key_variants:
            if variant in df_joop.columns:
                found_key_joop = variant
                break
        if found_key_joop:
            if found_key_joop != standard_key:
                df_joop.rename(columns={found_key_joop: standard_key}, inplace=True)
            df_joop[standard_key] = df_joop[standard_key].astype(str).str.strip().str.upper()
        else:
            st.warning("In het Joop Olieman-bestand is geen merge key gevonden.")
        
        # Pim K.I. Samen-bestand: verwacht exact de kolom "Stiercode NL / KI code"
        if "Stiercode NL / KI code" not in df_pim.columns:
            st.error("In het Pim K.I. Samen-bestand ontbreekt de kolom 'Stiercode NL / KI code'.")
            st.stop()
        else:
            df_pim.rename(columns={"Stiercode NL / KI code": standard_key}, inplace=True)
            df_pim[standard_key] = df_pim[standard_key].astype(str).str.strip().str.upper()
        
        # Prijslijst-bestand: verwacht exact de kolom "Artikelnr."
        if "Artikelnr." not in df_prijs.columns:
            st.error("In het Prijslijst-bestand ontbreekt de kolom 'Artikelnr.'.")
            st.stop()
        else:
            df_prijs.rename(columns={"Artikelnr.": standard_key}, inplace=True)
            df_prijs[standard_key] = df_prijs[standard_key].astype(str).str.strip().str.upper()
        
        # Debug: toon aantal overeenkomende sleutels tussen CRV en Pim
        match_count = len(set(df_crv[standard_key]) & set(df_pim[standard_key]))
        st.write("Aantal overeenkomende sleutels tussen CRV en Pim:", match_count)
        
        # Verwijder de kolommen uit CRV zodat de data uit Pim niet overschreven wordt
        cols_override = ["PFW", "aAa", "beta caseïne", "kappa Caseïne"]
        df_crv.drop(columns=cols_override, errors='ignore', inplace=True)
        
        # Wijzig de merge-volgorde: eerst CRV met Pim, daarna met Joop, daarna met Prijslijst
        df_merged = pd.merge(df_crv, df_pim, on=standard_key, how='left')
        
        if standard_key in df_joop.columns:
            # Voeg data uit Joop toe, met suffix indien nodig
            df_merged = pd.merge(df_merged, df_joop, on=standard_key, how='left', suffixes=("", "_joop"))
        else:
            st.warning("Merge key niet gevonden in het Joop Olieman-bestand.")
        
        if standard_key in df_prijs.columns:
            # Voeg data uit Prijslijst toe, met suffix indien nodig
            df_merged = pd.merge(df_merged, df_prijs, on=standard_key, how='left', suffixes=("", "_prijs"))
        else:
            st.warning("Merge key niet gevonden in het Prijslijst-bestand.")
        
        # Definitie van de gewenste kolomvolgorde en titels (volgens voorbeeld in tabblad 'fokstieren')
        kolom_mapping = [
            ("KI-code", "KI-code"),
            ("Stier", "Stiernaam"),
            ("Afstamming V", "Vader"),
            ("Afstamming MV", "M-vader"),
            ("PFW", "PFW"),
            ("aAa", "aAa"),
            ("beta caseïne", "beta caseïne"),
            ("kappa Caseïne", "kappa Caseïne"),
            ("Prijs", "Prijs"),
            ("Prijs gesekst", "Prijs gesekst"),
            ("% betrouwbaarheid", "Bt_1"),
            ("kg melk", "kgM"),
            ("% vet", "%V"),
            ("% eiwit", "%E"),
            ("kg vet", "kgV"),
            ("kg eiwit", "kgE"),
            ("INET", "INET"),
            ("NVI", "NVI"),
            ("TIP", "TIP"),
            ("% betrouwbaar", "Bt_5"),
            ("frame", "F"),
            ("uier", "U"),
            ("benen", "B_6"),
            ("totaal", "Ext"),
            ("hoogtemaat", "HT"),
            ("voorhand", "VH"),
            ("inhoud", "IH"),
            ("openheid", "OH"),
            ("conditie score", "CS"),
            ("kruisligging", "KL"),
            ("kruisbreedte", "KB"),
            ("beenstand achter", "BA"),
            ("beenstand zij", "BZ"),
            ("klauwhoek", "KH"),
            ("voorbeenstand", "VB"),
            ("beengebruik", "BG"),
            ("vooruieraanhechting", "VA"),
            ("voorspeenplaatsing", "VP"),
            ("speenlengte", "SL"),
            ("uierdiepte", "UD"),
            ("achteruierhoogte", "AH"),
            ("ophangband", "OB"),
            ("achterspeenplaatsing", "AP"),
            ("Geboortegemak", "Geb"),
            ("melksnelheid", "MS"),
            ("celgetal", "Cgt"),
            ("vruchtbaarheid", "Vru"),
            ("karakter", "KA"),
            ("laatrijpheid", "Ltrh"),
            ("Persistentie", "Pers"),
            ("klauwgezondheid", "Kgh"),
            ("levensduur", "Lvd"),
        ]
        
        # Bouw de finale DataFrame op basis van de mapping
        final_data = {}
        for final_header, source_column in kolom_mapping:
            if source_column in df_merged.columns:
                final_data[final_header] = df_merged[source_column]
            else:
                final_data[final_header] = None
        
        final_df = pd.DataFrame(final_data)
        
        # Exporteer de finale DataFrame naar een Excelbestand met de sheetnaam 'fokstieren'
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, sheet_name='fokstieren', index=False)
        output.seek(0)
        
        st.download_button(
            label="Download Stierenkaart Excel",
            data=output,
            file_name="stierenkaart.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Stierenkaart succesvol gegenereerd!")
    except Exception as e:
        st.error(f"Er is een fout opgetreden: {e}")
