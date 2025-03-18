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

# Functie om de merge key te bepalen
def bepaal_merge_key(df):
    # Mogelijke kolommen om op te matchen
    mogelijke_keys = ["KI-code", "levensnummer", "Stiernummer"]
    for key in mogelijke_keys:
        if key in df.columns:
            return key
    return None

if st.button("Genereer Stierenkaart"):
    if not (uploaded_crv and uploaded_joop and uploaded_pim and uploaded_prijs):
        st.warning("Zorg dat alle bestanden zijn geüpload!")
    else:
        try:
            # Lees de Excel-bestanden in (hier wordt ervan uitgegaan dat de relevante data in de eerste sheet staat)
            df_crv = pd.read_excel(uploaded_crv)
            df_joop = pd.read_excel(uploaded_joop)
            df_pim = pd.read_excel(uploaded_pim)
            df_prijs = pd.read_excel(uploaded_prijs)
            
            # Bepaal de merge key vanuit het CRV-bestand als uitgangspunt
            merge_key = bepaal_merge_key(df_crv)
            if merge_key is None:
                st.error("Geen geschikte merge key (bijv. 'KI-code' of 'levensnummer') gevonden in het CRV-bestand.")
            else:
                st.write(f"Gebruik merge key: **{merge_key}**")

                # Start met het CRV-bestand als basis en voeg de andere data eraan toe via een left-join
                df_merged = df_crv.copy()

                # Controleer of de merge key bestaat in de andere DataFrames, en voeg ze samen
                if merge_key in df_joop.columns:
                    df_merged = pd.merge(df_merged, df_joop, on=merge_key, how='left')
                else:
                    st.warning(f"Merge key '{merge_key}' niet gevonden in het Joop Olieman-bestand.")

                if merge_key in df_pim.columns:
                    df_merged = pd.merge(df_merged, df_pim, on=merge_key, how='left')
                else:
                    st.warning(f"Merge key '{merge_key}' niet gevonden in het Pim K.I. Samen-bestand.")

                if merge_key in df_prijs.columns:
                    df_merged = pd.merge(df_merged, df_prijs, on=merge_key, how='left')
                else:
                    st.warning(f"Merge key '{merge_key}' niet gevonden in het Prijslijst-bestand.")

                # Hier kun je desgewenst kolommen hernoemen of de volgorde aanpassen volgens de mapping:
                # Bijvoorbeeld: df_merged.rename(columns={'Stiernaam': 'Stier'}, inplace=True)
                # Voeg hier de overige gewenste bewerkingen toe volgens jouw specificaties.

                # Genereer de Excel als een in-memory bestand met de sheetnaam 'fokstieren'
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
