import streamlit as st
import pandas as pd

# Titel van de app
st.title("🐂 Stieren Data Selectie op Naam")

# Upload een Excel-bestand
uploaded_file = st.file_uploader("📂 Upload je Excel-bestand", type=["xlsx"])

# Controleer of er een bestand is geüpload
if uploaded_file is not None:
    try:
        # Laad het Excel-bestand in een Pandas DataFrame, met rij 2 als kolomnamen
        df = pd.read_excel(uploaded_file, engine="openpyxl", header=1)  # Rij 2 als header (index=1)

        # Laat een voorbeeld van de data zien
        st.write("📊 **Voorbeeld van de data:**")
        st.dataframe(df.head())

        # **Stieren staan altijd in kolom C (index 2, want Python begint bij 0)**
        stieren_kolom = df.columns[2]  # Kolom C vastzetten

        # Haal unieke stierennamen op
        stieren_namen = df[stieren_kolom].dropna().unique()

        # Multiselectie van stierennamen
        geselecteerde_stieren = st.multiselect("🐂 Selecteer de stieren", stieren_namen)

        if geselecteerde_stieren:
            # Filter de data op de geselecteerde stierennamen
            gefilterde_data = df[df[stieren_kolom].isin(geselecteerde_stieren)]

            # **Gebruiker kiest de kolomvolgorde**
            st.write("📌 **Sleep de kolommen in de gewenste volgorde:**")
            kolommen = list(df.columns)
            geselecteerde_kolommen = st.multiselect("📋 Kies kolommen", kolommen, default=kolommen)

            if geselecteerde_kolommen:
                gefilterde_data = gefilterde_data[geselecteerde_kolommen]  # Alleen gekozen kolommen tonen

            # **Sorteeropties op basis van kolomnamen**
            sorteer_keuze = st.selectbox("🔽 Sorteer op kolom:", geselecteerde_kolommen, index=0)

            # **Sorteer de gefilterde data**
            gesorteerde_data = gefilterde_data.sort_values(by=sorteer_keuze)

            # **Laat de gesorteerde data zien**
            st.write("✅ **Gesorteerde Data:**")
            st.dataframe(gesorteerde_data)

            # **Download-knop voor de gesorteerde data**
            csv = gesorteerde_data.to_csv(index=False).encode('utf-8')
            st.download_button(label="⬇️ Download CSV", data=csv, file_name="gesorteerde_data.csv", mime="text/csv")

    except Exception as e:
        st.error(f"❌ Er is een fout opgetreden bij het verwerken van het bestand: {e}")

else:
    st.warning("⚠️ Upload een Excel-bestand om te beginnen.")
