import streamlit as st
import pandas as pd
import io

st.title("Stierenkaart Generator")
st.write("Upload de volgende bestanden om de stierenkaart te genereren:")

# Bestandsuploaders voor de verschillende bronbestanden
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
            
            # Gebruik als basis de CRV file en verwacht dat deze de kolom "KI-code" bevat
            if "KI-code" in df_crv.columns:
                base_key = "KI-code"
            else:
                st.error("De kolom 'KI-code' is niet aanwezig in het CRV-bestand. Zorg voor consistente data.")
                st.stop()
            
            st.write(f"Gebruik merge key: **{base_key}**")
            
            # In het Joop Olieman bestand verwachten we al de kolom "KI-code"
            if base_key not in df_joop.columns:
                st.warning(f"In het Joop Olieman bestand ontbreekt de kolom '{base_key}'.")
            
            # In het Pim K.I. Samen bestand hernoemen we 'stiecode' naar 'KI-code'
            if "stiecode" in df_pim.columns:
                df_pim.rename(columns={"stiecode": base_key}, inplace=True)
            else:
                st.warning("In het Pim K.I. Samen bestand ontbreekt de kolom 'stiecode'.")
            
            # In het Prijslijst bestand hernoemen we 'artikelnummer' naar 'KI-code'
            if "artikelnummer" in df_prijs.columns:
                df_prijs.rename(columns={"artikelnummer": base_key}, inplace=True)
            else:
                st.warning("In het Prijslijst bestand ontbreekt de kolom 'artikelnummer'.")
            
            # Start met het CRV-bestand als basis en voeg de andere data eraan toe via left-joins
            df_merged = df_crv.copy()
            
            if base_key in df_joop.columns:
                df_merged = pd.merge(df_merged, df_joop, on=base_key, how='left')
            else:
                st.warning(f"Merge key '{base_key}' niet gevonden in het Joop Olieman bestand.")
            
            if base_key in df_pim.columns:
                df_merged = pd.merge(df_merged, df_pim, on=base_key, how='left')
            else:
                st.warning(f"Merge key '{base_key}' niet gevonden in het Pim K.I. Samen bestand.")
            
            if base_key in df_prijs.columns:
                df_merged = pd.merge(df_merged, df_prijs, on=base_key, how='left')
            else:
                st.warning(f"Merge key '{base_key}' niet gevonden in het Prijslijst bestand.")
            
            # Genereer een Excel bestand met de sheetnaam 'fokstieren'
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
