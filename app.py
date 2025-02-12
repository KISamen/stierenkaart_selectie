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

        # Laad het extra Excel-bestand met stierinformatie
        df_extra = pd.read_excel(extra_file, engine="openpyxl", header=1)

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

            # **Kolommen uit het extra bestand ophalen**
            kappa_caseine_kolom = df_extra.columns[15]  # Kolom P (index 15)
            beta_caseine_kolom = df_extra.columns[16]  # Kolom Q (index 16)

            # Stieren matchen op Stiernaam, Levensnummer of KI-code
            merge_keys = [stieren_kolom]  # Start met matchen op Stiernaam
            if "Levensnummer" in df.columns and "Levensnummer" in df_extra.columns:
                merge_keys.append("Levensnummer")
            if "Kicode" in df.columns and "Kicode" in df_extra.columns:
                merge_keys.append("Kicode")

            # Merge de gegevens
            gefilterde_data = gefilterde_data.merge(df_extra[[stieren_kolom, kappa_caseine_kolom, beta_caseine_kolom]],
                                                   on=merge_keys, how="left")

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
                    "Vader": "Father",  # Vader toegevoegd
                    "Moeders Vader": "Maternal Grandfather",  # Moeders Vader toegevoegd
                    "aAa": "aAa", "% Betr": "% reliability", "Kg melk": "kg milk",
                    "% vet": "% fat", "% eiwit": "% protein", "Kg vet": "kg fat",
                    "Kg eiwit": "kg protein", "Dcht totaal": "#Daughters", "% Betr.1": "% reliability",
                    "Frame": "frame", "Uier": "udder", "Beenwerk": "feet & legs",
                    "Totaal exterieur": "final score", "Hoogtemaat": "stature", "Voorhand": "chest width",
                    "Inhoud": "body depth", "Openheid": "angularity", "Conditie score": "condition score",
                    "Kruisligging": "rump angle", "Kruisbreedte": "rump width", "Beenstand achter": "rear legs rear view",
                    "Beenstand zij": "rear leg set", "Klauwhoek": "foot angle", "Voorbeenstand": "front feet orientation",
                    "Beengebruik": "mobility", "Vooruieraanhechting": "fore udder attachment", 
                    "Voorspeenplaatsing": "front teat placement", "Speenlengte": "teat length",
                    "Uierdiepte": "udder depth", "Achteruierhoogte": "rear udder height", "Ophangband": "central ligament",
                    "Achterspeenplaatsing": "rear teat placement", "Uierbalans": "udder balance",
                    "Geboortegemak": "calving ease", "Melksnelheid": "milking speed", "Celgetal": "somatic cell score",
                    "Vruchtbaarheid": "female fertility", "Karakter": "temperament", 
                    "Verwantschapsgraad": "maturity rate", "Persistentie": "persistence",
                    "Klauwgezondheid": "hoof health",
                    kappa_caseine_kolom: "Kappa-caseine",
                    beta_caseine_kolom: "Beta-caseine"
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
        st.error(f"❌ Er is een fout opgetreden bij het verwerken van het bestand: {e}")

else:
    st.warning("⚠️ Upload beide Excel-bestanden om te beginnen.")
