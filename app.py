import streamlit as st
import pandas as pd
import io

# Titel van de app
st.title("🐂 Stieren Data Selectie")

# Upload een Excel-bestand
uploaded_file = st.file_uploader("📂 Upload je Excel-bestand", type=["xlsx"])

# Controleer of er een bestand is geüpload
if uploaded_file is not None:
    try:
        # Laad het Excel-bestand in een Pandas DataFrame, met rij 2 als kolomnamen
        df = pd.read_excel(uploaded_file, engine="openpyxl", header=1)

        # Laat een voorbeeld van de data zien
        st.write("📊 **Voorbeeld van de data:**")
        st.dataframe(df.head())

        # **Stieren staan altijd in kolom C (index 2, want Python begint bij 0)**
        stieren_kolom = df.columns[2]  # Kolom C vastzetten
        ki_kolom = df.columns[1]  # Kolom B bevat de KI-code

        # Haal unieke stierennamen op
        stieren_namen = df[stieren_kolom].dropna().unique()

        # Multiselectie van stierennamen
        geselecteerde_stieren = st.multiselect("🐂 Selecteer de stieren", stieren_namen)

        if geselecteerde_stieren:
            # Filter de data op de geselecteerde stierennamen
            gefilterde_data = df[df[stieren_kolom].isin(geselecteerde_stieren)]

            # **Stap 1: Keuze tussen Drag & Drop en Canada Template**
            optie = st.radio("📌 Kies een optie:", ["📂 Drag & Drop kolommen", "🇨🇦 Canada-template"])

            if optie == "📂 Drag & Drop kolommen":
                st.write("📌 **Sleep de kolommen in de gewenste volgorde:**")
                kolommen = list(df.columns)
                geselecteerde_kolommen = st.multiselect("📋 Kies kolommen", kolommen, default=kolommen)
            
            else:
                # **Canada-template volgorde instellen (met correcte kolomnamen)**
                canada_volgorde = {
                    "Kicode": "KI Code",  # KI-code toegevoegd
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
                    "Klauwgezondheid": "hoof health"
                }

                # Alleen kolommen gebruiken die in de dataset staan
                geselecteerde_kolommen = [col for col in canada_volgorde.keys() if col in df.columns]

                # **Stap 2: Uitklapbaar veld voor vertalingen**
                vertalingen = {}
                with st.expander("🌍 **Vertaling aanpassen**"):
                    for col in geselecteerde_kolommen:
                        vertalingen[col] = st.text_input(f"Vertaling voor '{col}':", value=canada_volgorde[col])

                # Pas de vertalingen toe
                gefilterde_data = gefilterde_data[geselecteerde_kolommen]
                gefilterde_data = gefilterde_data.rename(columns=vertalingen)

            if geselecteerde_kolommen:
                # **Sorteeropties op basis van kolomnamen**
                sorteer_keuze = st.selectbox("🔽 Sorteer op kolom:", list(gefilterde_data.columns), index=0)

                # **Sorteer de gefilterde data**
                gesorteerde_data = gefilterde_data.sort_values(by=sorteer_keuze)

                # **Laat de gesorteerde data zien**
                st.write("✅ **Gesorteerde Data:**")
                st.dataframe(gesorteerde_data)

                # **Output als Excel (.xlsx)**
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    gesorteerde_data.to_excel(writer, index=False, sheet_name="Data")
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
    st.warning("⚠️ Upload een Excel-bestand om te beginnen.")
