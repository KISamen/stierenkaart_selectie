import streamlit as st
import pandas as pd
import io

# Titel van de app
st.title("🐂 Stieren Data Selectie")

# Upload het hoofd Excel-bestand
uploaded_file = st.file_uploader("📂 Upload je Excel-bestand", type=["xlsx"], key="file1")

# Upload het extra bestand met Kappa-caseïne en Beta-caseïne
extra_file = st.file_uploader("📂 Upload extra stierinfo-bestand", type=["xlsx"], key="file2")

# Controleer of beide bestanden zijn geüpload
if uploaded_file is not None and extra_file is not None:
    try:
        # Laad het hoofd Excel-bestand in een Pandas DataFrame, met rij 2 als kolomnamen
        df = pd.read_excel(uploaded_file, engine="openpyxl", header=1)  # Rij 2 als header (index=1)
        
        # Standaardiseer kolomnamen in het hoofdbestand
        df.rename(columns={
            "Levensnummer (ITB)": "Levensnummer",
            "KI-code": "Kicode"
        }, inplace=True)

        # Laad het extra Excel-bestand met de juiste header-rij
        df_extra = pd.read_excel(extra_file, engine="openpyxl", header=0)
        df_extra.columns = df_extra.iloc[0].astype(str)  # Converteer kolomnamen naar string
        df_extra = df_extra[1:].reset_index(drop=True)  # Verwijder de duplicaat-header rij
        
        # Verwijder ongewenste 'Unnamed' kolommen (om fouten te voorkomen)
        df_extra = df_extra.loc[:, df_extra.columns.notna() & ~df_extra.columns.astype(str).str.contains('^Unnamed')]

        # Controleer dubbele kolommen en voeg achtervoegsel toe
        duplicate_cols = [col for col in df_extra.columns if col in df.columns]
        df_extra.rename(columns={col: f"{col}_extra" for col in duplicate_cols}, inplace=True)

        # Herbenoem relevante kolommen in het extra bestand
        df_extra.rename(columns={
            "Naam": "Stiernaam",
            "Levensnummer": "Levensnummer",
            "KI-code": "Kicode",
            "Kappa-caseine": "Kappa-caseine",
            "Betacasine": "Beta-caseine"
        }, inplace=True)

        # Laat een voorbeeld van de data zien
        st.write("📊 **Voorbeeld van de data:**")
        st.dataframe(df.head())

        # Haal unieke stierennamen op
        stieren_namen = df["Stiernaam"].dropna().unique()

        # Multiselectie van stierennamen
        geselecteerde_stieren = st.multiselect("🐂 Selecteer de stieren", stieren_namen)

        if geselecteerde_stieren:
            # Filter de data op de geselecteerde stierennamen
            gefilterde_data = df[df["Stiernaam"].isin(geselecteerde_stieren)]

            # Merge-sleutels bepalen
            merge_keys = ["Stiernaam", "Levensnummer", "Kicode"]
            merge_keys = [key for key in merge_keys if key in df.columns and key in df_extra.columns]

            # Merge de gegevens met extra bestand
            gefilterde_data = gefilterde_data.merge(
                df_extra[["Stiernaam", "Levensnummer", "Kicode", "Kappa-caseine", "Beta-caseine"]],
                on=merge_keys,
                how="left"
            )

            # **Stap 1: Keuze tussen Drag & Drop en Canada Template**
            optie = st.radio("📌 Kies een optie:", ["📂 Drag & Drop kolommen", "🇨🇦 Canada-template"])

            if optie == "📂 Drag & Drop kolommen":
                st.write("📌 **Sleep de kolommen in de gewenste volgorde:**")
                kolommen = list(gefilterde_data.columns)
                geselecteerde_kolommen = st.multiselect("📋 Kies kolommen", kolommen, default=kolommen)
            else:
                # **Canada-template volgorde instellen (met correcte kolomnamen)**
                canada_volgorde = {
                    "Stiernaam": "Bull name",
                    "Kappa-caseine": "Kappa-caseine",
                    "Beta-caseine": "Beta-caseine"
                }

                geselecteerde_kolommen = [col for col in canada_volgorde.keys() if col in gefilterde_data.columns]
                gefilterde_data = gefilterde_data[geselecteerde_kolommen]
                gefilterde_data = gefilterde_data.rename(columns=canada_volgorde)

            # **Output als Excel (.xlsx)**
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                gefilterde_data.to_excel(writer, index=False, sheet_name="Data")
            output.seek(0)

            # **Download-knop voor de Excel-output**
            st.download_button(
                label="⬇️ Download Excel",
                data=output,
                file_name="gesorteerde_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"❌ Er is een fout opgetreden bij het verwerken van de bestanden: {e}")
else:
    st.warning("⚠️ Upload beide Excel-bestanden om te beginnen.")
