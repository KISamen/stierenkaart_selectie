import streamlit as st
import pandas as pd
import io

st.title("Stierenkaart Generator")
st.write("Upload de volgende bestanden om de stierenkaart te genereren:")

uploaded_crv = st.file_uploader("Upload Bronbestand CRV (bijv. 'Bronbestand CRV DEC2024.xlsx')", type=["xlsx"])
uploaded_joop = st.file_uploader("Upload Bronbestand Joop Olieman (bijv. 'Bronbestand Joop Olieman.xlsx')", type=["xlsx"])
uploaded_pim = st.file_uploader("Upload Pim K.I. Samen (bijv. 'Pim K.I. Samen.xlsx')", type=["xlsx"])
uploaded_prijs = st.file_uploader("Upload Prijslijst (bijv. 'Prijslijst.xlsx')", type=["xlsx"])

if st.button("Genereer Stierenkaart"):
    if not (uploaded_crv and uploaded_joop and uploaded_pim and uploaded_prijs):
        st.warning("Zorg dat alle bestanden zijn geüpload!")
    else:
        try:
            # Lees de Excel-bestanden in
            df_crv = pd.read_excel(uploaded_crv)
            df_joop = pd.read_excel(uploaded_joop)
            df_pim = pd.read_excel(uploaded_pim)
            df_prijs = pd.read_excel(uploaded_prijs)
            
            # Stel een standaard merge key in
            standard_key = "KI-code"
            key_variants = ["KI-code", "KI code", "KI-Code", "ki code"]
            
            # Zoek naar de merge key in het CRV-bestand en hernoem naar de standaard key
            found_key = None
            for variant in key_variants:
                if variant in df_crv.columns:
                    found_key = variant
                    break
            if not found_key:
                st.error("De kolom 'KI-code' (of een variant) is niet aanwezig in het CRV-bestand. Zorg voor consistente data.")
                st.stop()
            if found_key != standard_key:
                df_crv.rename(columns={found_key: standard_key}, inplace=True)
            
            st.write(f"Gebruik merge key: **{standard_key}**")
            
            # Voor het Joop Olieman-bestand: zoek en hernoem de merge key indien aanwezig
            found_key_joop = None
            for variant in key_variants:
                if variant in df_joop.columns:
                    found_key_joop = variant
                    break
            if found_key_joop:
                if found_key_joop != standard_key:
                    df_joop.rename(columns={found_key_joop: standard_key}, inplace=True)
            else:
                st.warning(f"In het Joop Olieman-bestand is geen kolom gevonden die overeenkomt met '{standard_key}'.")
            
            # Voor het Pim K.I. Samen-bestand: hernoem 'stiecode' naar de standaard merge key
            if "stiecode" in df_pim.columns:
                df_pim.rename(columns={"stiecode": standard_key}, inplace=True)
            else:
                st.warning("In het Pim K.I. Samen-bestand ontbreekt de kolom 'stiecode'.")
            
            # Voor het Prijslijst-bestand: hernoem 'artikelnummer' naar de standaard merge key
            if "artikelnummer" in df_prijs.columns:
                df_prijs.rename(columns={"artikelnummer": standard_key}, inplace=True)
            else:
                st.warning("In het Prijslijst-bestand ontbreekt de kolom 'artikelnummer'.")
            
            # Start met het CRV-bestand als basis en voeg de andere data toe via left-joins
            df_merged = df_crv.copy()
            
            if standard_key in df_joop.columns:
                df_merged = pd.merge(df_merged, df_joop, on=standard_key, how='left')
            else:
                st.warning(f"Merge key '{standard_key}' niet gevonden in het Joop Olieman-bestand.")
            
            if standard_key in df_pim.columns:
                df_merged = pd.merge(df_merged, df_pim, on=standard_key, how='left')
            else:
                st.warning(f"Merge key '{standard_key}' niet gevonden in het Pim K.I. Samen-bestand.")
            
            if standard_key in df_prijs.columns:
                df_merged = pd.merge(df_merged, df_prijs, on=standard_key, how='left')
            else:
                st.warning(f"Merge key '{standard_key}' niet gevonden in het Prijslijst-bestand.")
            
            # Genereer een Excelbestand met de sheetnaam 'fokstieren'
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_merged.to_excel(writer, sheet_name='fokstieren', index=False)
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
