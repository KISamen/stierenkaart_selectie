import streamlit as st
import pandas as pd

# Titel van de app
st.title("🐂 Stieren Data Selectie")

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

            # **Stap 1: Keuze tussen Drag & Drop en Canada Template**
            optie = st.radio("📌 Kies een optie:", ["📂 Drag & Drop kolommen", "🇨🇦 Canada-template"])

            if optie == "📂 Drag & Drop kolommen":
                # **Gebruiker kiest de kolomvolgorde**
                st.write("📌 **Sleep de kolommen in de gewenste volgorde:**")
                kolommen = list(df.columns)
                geselecteerde_kolommen = st.multiselect("📋 Kies kolommen", kolommen, default=kolommen)
            
            else:
                # **Canada-template volgorde instellen (met correcte kolomnamen)**
                canada_volgorde = {
                    "Stiernaam": "Bull name",
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
                    "  Geboortegemak": "calving ease",  # BR kolom
                    "Melksnelheid": "milking speed", "Celgetal": "somatic cell score",
                    "Vruchtbaarheid": "female fertility", "Karakter": "temperament", 
                    "Verwantschapsgraad": "maturity rate", "Persistentie": "persistence",
                    "Klauwgezondheid": "hoof health"
                }

                # Alleen kolommen gebruiken die in de dataset staan
                geselecteerde_kolommen = [col for col in canada_volgorde.keys() if col in df.columns]

                # **Stap 2: Optie om vertalingen aan te passen**
                st.write("🌍 **Pas vertalingen aan (optioneel):**")
                vertalingen = {}
                for col in geselecteerde_kolommen:
                    vertalingen[col] = st.text_input(f"Vertaling voor '{col}':", value=canada_volgorde[col])

                # **Stap 3: Controle op dubbele namen en oplossen**
                def maak_unieke_namen(naam_lijst):
                    unieke_namen = {}
                    nieuwe_namen = []
                    for naam in naam_lijst:
                        if naam in unieke_namen:
                            unieke_namen[naam] += 1
                            nieuwe_namen.append(f"{naam}_{unieke_namen[naam]}")
                        else:
                            unieke_namen[naam] = 0
                            nieuwe_namen.append(naam)
                    return nieuwe_namen

                # Pas de vertalingen toe en maak namen uniek indien nodig
                nieuwe_kolomnamen = [vertalingen[col] for col in geselecteerde_kolommen]
                unieke_kolomnamen = maak_unieke_namen(nieuwe_kolomnamen)

                # Hernoem de kolommen in de dataset
                gefilterde_data = gefilterde_data[geselecteerde_kolommen]
                gefilterde_data.columns = unieke_kolomnamen

            if geselecteerde_kolommen:
                # **Sorteeropties op basis van kolomnamen**
                sorteer_keuze = st.selectbox("🔽 Sorteer op kolom:", list(gefilterde_data.columns), index=0)

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
