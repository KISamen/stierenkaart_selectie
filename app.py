import streamlit as st
import pandas as pd

# Titel van de app
st.title("Stieren Data Selectie")

# Upload een Excel-bestand
uploaded_file = st.file_uploader("Upload je Excel-bestand", type=["xlsx"])

# Controleer of er een bestand is geüpload
if uploaded_file is not None:
    try:
        # Laad het Excel-bestand in een Pandas DataFrame
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        
        # Laat een voorbeeld van de data zien
        st.write("📊 **Voorbeeld van de data:**")
        st.dataframe(df.head())

        # Selecteer stieren (eerste kolom wordt aangenomen als stierenlijst)
        if not df.empty:
            stieren = df.iloc[:, 0].dropna().unique()  # Verander de index als de kolom een andere positie heeft
            geselecteerde_stieren = st.multiselect("🔍 Selecteer de stieren", stieren)

            if geselecteerde_stieren:
                # Filter de data op de geselecteerde stieren
                gefilterde_data = df[df.iloc[:, 0].isin(geselecteerde_stieren)]

                # Sorteeropties op basis van kolomnamen
                sorteer_opties = list(df.columns)
                sorteer_keuze = st.selectbox("🔽 Sorteer op kolom:", sorteer_opties, index=1)

                # Sorteer de gefilterde data
                gesorteerde_data = gefilterde_data.sort_values(by=sorteer_keuze)

                # Laat de gesorteerde data zien
                st.write("✅ **Gesorteerde Data:**")
                st.dataframe(gesorteerde_data)

                # Download-knop voor de gesorteerde data
                csv = gesorteerde_data.to_csv(index=False).encode('utf-8')
                st.download_button(label="⬇️ Download CSV", data=csv, file_name="gesorteerde_data.csv", mime="text/csv")
    
    except Exception as e:
        st.error(f"❌ Er is een fout opgetreden bij het verwerken van het bestand: {e}")

else:
    st.warning("⚠️ Upload een Excel-bestand om te beginnen.")

