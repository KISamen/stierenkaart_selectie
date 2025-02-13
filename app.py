import streamlit as st
import pandas as pd
import io

# Titel van de app
st.title("🐂 Stieren Data Selectie")

# Upload de twee Excel-bestanden
uploaded_tip_file = st.file_uploader("📂 Upload je TIP Excel-bestand", type=["xlsx"])
uploaded_stierinfo_file = st.file_uploader("📂 Upload het Stierinfo Excel-bestand", type=["xlsx"])

if uploaded_tip_file is not None and uploaded_stierinfo_file is not None:
    try:
        # Laad de bestanden in Pandas DataFrames
        tip_df = pd.read_excel(uploaded_tip_file, engine="openpyxl", header=1)
        stierinfo_df = pd.read_excel(uploaded_stierinfo_file, engine="openpyxl")

        # Hernoem dubbele kolommen
        tip_df = tip_df.loc[:, ~tip_df.columns.duplicated()]
        stierinfo_df = stierinfo_df.loc[:, ~stierinfo_df.columns.duplicated()]

        # Relevante kolommen uit TIP-bestand selecteren
        tip_df = tip_df.rename(columns={
            "Kicode": "KI Code",
            "Naam": "Stiernaam"
        })

        # Relevante kolommen uit Stierinfo-bestand selecteren
        stierinfo_df = stierinfo_df.rename(columns={
            "Levensnummer": "Stier ID",
            "Kappa-caseine": "Kappa Caseïne",
            "Betacasine": "Beta Caseïne"
        })

        # Zorg ervoor dat de merge kolommen dezelfde datatype hebben
        tip_df["KI Code"] = tip_df["KI Code"].astype(str)
        stierinfo_df["Stier ID"] = stierinfo_df["Stier ID"].astype(str)
        
        # Merge de data op basis van stiernaam of KI-code
        merged_df = tip_df.merge(stierinfo_df, left_on="KI Code", right_on="Stier ID", how="left")
        
        # Laat een voorbeeld van de data zien
        st.write("📊 **Voorbeeld van de data na samenvoeging:**")
        st.dataframe(merged_df.head())

        # Multiselectie van stierennamen
        stieren_namen = merged_df["Stiernaam"].dropna().unique()
        geselecteerde_stieren = st.multiselect("🐂 Selecteer de stieren", stieren_namen)

        if geselecteerde_stieren:
            # Filter de data op de geselecteerde stierennamen
            gefilterde_data = merged_df[merged_df["Stiernaam"].isin(geselecteerde_stieren)]

            # **Canada-template volgorde instellen (met correcte kolomnamen)**
            canada_volgorde = {
                "KI Code": "KI Code",
                "Stiernaam": "Bull name",
                "Vader": "Father",
                "Moeders Vader": "Maternal Grandfather",
                "aAa": "aAa",
                "% Betr": "% reliability",
                "Kg melk": "kg milk",
                "% vet": "% fat",
                "% eiwit": "% protein",
                "Kg vet": "kg fat",
                "Kg eiwit": "kg protein",
                "Dcht totaal": "#Daughters",
                "% Betr.1": "% reliability",
                "Frame": "frame",
                "Uier": "udder",
                "Beenwerk": "feet & legs",
                "Totaal exterieur": "final score",
                "Hoogtemaat": "stature",
                "Voorhand": "chest width",
                "Inhoud": "body depth",
                "Openheid": "angularity",
                "Conditie score": "condition score",
                "Kruisligging": "rump angle",
                "Kruisbreedte": "rump width",
                "Beenstand achter": "rear legs rear view",
                "Beenstand zij": "rear leg set",
                "Klauwhoek": "foot angle",
                "Voorbeenstand": "front feet orientation",
                "Beengebruik": "mobility",
                "Vooruieraanhechting": "fore udder attachment",
                "Voorspeenplaatsing": "front teat placement",
                "Speenlengte": "teat length",
                "Uierdiepte": "udder depth",
                "Achteruierhoogte": "rear udder height",
                "Ophangband": "central ligament",
                "Achterspeenplaatsing": "rear teat placement",
                "Uierbalans": "udder balance",
                "Geboortegemak": "calving ease",
                "Melksnelheid": "milking speed",
                "Celgetal": "somatic cell score",
                "Vruchtbaarheid": "female fertility",
                "Karakter": "temperament",
                "Verwantschapsgraad": "maturity rate",
                "Persistentie": "persistence",
                "Klauwgezondheid": "hoof health",
                "Kappa Caseïne": "Kappa Casein",
                "Beta Caseïne": "Beta Casein"
            }

            # Alleen kolommen gebruiken die in de dataset staan
            geselecteerde_kolommen = [col for col in canada_volgorde.keys() if col in gefilterde_data.columns]

            # **Uitklapbaar veld voor vertalingen**
            vertalingen = {}
            with st.expander("🌍 **Vertaling aanpassen**"):
                for col in geselecteerde_kolommen:
                    vertalingen[col] = st.text_input(f"Vertaling voor '{col}':", value=canada_volgorde[col])

            # Pas de vertalingen toe
            gefilterde_data = gefilterde_data[geselecteerde_kolommen]
            gefilterde_data.columns = [vertalingen[col] for col in geselecteerde_kolommen]

            # **Laat de gefilterde data zien**
            st.write("✅ **Gefilterde Data:**")
            st.dataframe(gefilterde_data)

            # **Output als Excel (.xlsx)**
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                gefilterde_data.to_excel(writer, index=False, sheet_name="Data")
            output.seek(0)

            # **Download-knop voor de Excel-output**
            st.download_button(
                label="⬇️ Download Excel",
                data=output,
                file_name="stieren_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"❌ Er is een fout opgetreden bij het verwerken van de bestanden: {e}")
else:
    st.warning("⚠️ Upload beide Excel-bestanden om te beginnen.")
