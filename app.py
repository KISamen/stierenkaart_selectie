import streamlit as st
import pandas as pd
import io

# Titel van de app
st.title("🐂 Stieren Data Selectie")

# Upload het hoofd Excel-bestand
uploaded_file = st.file_uploader("📂 Upload je hoofd Excel-bestand", type=["xlsx"], key="file1")

# Upload het extra bestand met Kappa-caseïne en Beta-caseïne
extra_file = st.file_uploader("📂 Upload extra stierinfo-bestand", type=["xlsx"], key="file2")

# Controleer of beide bestanden zijn geüpload
if uploaded_file is not None and extra_file is not None:
    try:
        # Laad het hoofd Excel-bestand in een Pandas DataFrame, met rij 2 als kolomnamen
        df = pd.read_excel(uploaded_file, engine="openpyxl", header=1)  # Pas 'header' aan indien nodig

        # Laad het extra Excel-bestand met stierinformatie
        df_extra = pd.read_excel(extra_file, engine="openpyxl", header=1)  # Pas 'header' aan indien nodig

        # Toon de kolomnamen van beide DataFrames voor controle
        st.write("Kolomnamen in het hoofd DataFrame:", df.columns.tolist())
        st.write("Kolomnamen in het extra DataFrame:", df_extra.columns.tolist())

        # Controleer of de benodigde kolommen aanwezig zijn
        required_columns = ['Stiernaam', 'Levensnummer', 'Kicode']
        for col in required_columns:
            if col not in df.columns:
                st.error(f"❌ Kolom '{col}' ontbreekt in het hoofd DataFrame.")
            if col not in df_extra.columns:
                st.error(f"❌ Kolom '{col}' ontbreekt in het extra DataFrame.")

        # Ga verder met de verwerking als alle benodigde kolommen aanwezig zijn
        if all(col in df.columns for col in required_columns) and all(col in df_extra.columns for col in required_columns):
            # Haal unieke stierennamen op
            stieren_namen = df['Stiernaam'].dropna().unique()

            # Multiselectie van stierennamen
            geselecteerde_stieren = st.multiselect("🐂 Selecteer de stieren", stieren_namen)

            if geselecteerde_stieren:
                # Filter de data op de geselecteerde stierennamen
                gefilterde_data = df[df['Stiernaam'].isin(geselecteerde_stieren)]

                # Kolommen uit het extra bestand ophalen
                kappa_caseine_kolom = 'Kappa-caseine'  # Pas aan indien de kolom een andere naam heeft
                beta_caseine_kolom = 'Beta-caseine'    # Pas aan indien de kolom een andere naam heeft

                # Merge de gegevens op basis van 'Stiernaam', 'Levensnummer' en 'Kicode'
                merge_keys = ['Stiernaam', 'Levensnummer', 'Kicode']
                gefilterde_data = gefilterde_data.merge(
                    df_extra[[*merge_keys, kappa_caseine_kolom, beta_caseine_kolom]],
                    on=merge_keys,
                    how="left"
                )

                # Verdere verwerking en output...

    except Exception as e:
        st.error(f"❌ Er is een fout opgetreden bij het verwerken van de bestanden: {e}")

else:
    st.warning("⚠️ Upload beide Excel-bestanden om te beginnen.")
