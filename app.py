import streamlit as st
import pandas as pd
import io

st.title("Stierenkaart Generator")

st.markdown("""
Upload de volgende bestanden:
- **Bronbestand CRV DEC2024.xlsx** (KI-Code in kolom A)
- **PIM K.I. Samen.xlsx** (Stiercode NL / KI code in kolom M)
- **Prijslijst.xlsx** (Artikelnr. als KI-code)
- **Bronbestand Joop Olieman.xlsx** (Kicode in kolom B)
""")

# Bestanden uploaden
uploaded_crv = st.file_uploader("Upload Bronbestand CRV DEC2024.xlsx", type=["xlsx"])
uploaded_pim = st.file_uploader("Upload PIM K.I. Samen.xlsx", type=["xlsx"])
uploaded_prijslijst = st.file_uploader("Upload Prijslijst.xlsx", type=["xlsx"])
uploaded_joop = st.file_uploader("Upload Bronbestand Joop Olieman.xlsx", type=["xlsx"])

def load_excel(file):
    try:
        return pd.read_excel(file)
    except Exception as e:
        st.error(f"Fout bij het laden van het bestand: {e}")
        return None

if st.button("Genereer Stierenkaart"):
    if not (uploaded_crv and uploaded_pim and uploaded_prijslijst and uploaded_joop):
        st.error("Zorg dat je alle vereiste bestanden uploadt!")
    else:
        # Lees de bestanden
        df_crv = load_excel(uploaded_crv)
        df_pim = load_excel(uploaded_pim)
        df_prijslijst = load_excel(uploaded_prijslijst)
        df_joop = load_excel(uploaded_joop)

        # Gebruik expliciete check op None
        if any(df is None for df in [df_crv, df_pim, df_prijslijst, df_joop]):
            st.error("Er is een fout opgetreden bij het laden van een of meerdere bestanden.")
        else:
            # Hernoem de kolommen zodat de merge key overal "KI_Code" heet
            df_crv.rename(columns={"KI-Code": "KI_Code"}, inplace=True)
            df_pim.rename(columns={"Stiercode NL / KI code": "KI_Code"}, inplace=True)
            df_prijslijst.rename(columns={"Artikelnr.": "KI_Code"}, inplace=True)
            df_joop.rename(columns={"Kicode": "KI_Code"}, inplace=True)

            # Eerste merge: CRV als basis, voeg PIM toe
            df_merged = pd.merge(df_crv, df_pim, on="KI_Code", how="left", suffixes=("", "_pim"))
            # Voeg prijslijst toe
            df_merged = pd.merge(df_merged, df_prijslijst, on="KI_Code", how="left", suffixes=("", "_prijslijst"))
            # Voeg Joop Olieman toe
            df_merged = pd.merge(df_merged, df_joop, on="KI_Code", how="left", suffixes=("", "_joop"))

            # Definieer de kolommapping voor de stierenkaart
            kolom_mapping = {
                "KI-Code": "KI_Code",
                "Eigenaarscode": "Eigenaarscode",        # Pas aan indien de kolomnaam anders is
                "Stiernummer": "Stiernummer",            # idem
                "Stiernaam": "Stier",                    # uit CRV (bijvoorbeeld)
                "Erf-fact": "Erf-fact",                  # idem
                "Vader": "Afstamming V",                 # uit CRV
                "M-vader": "Afstamming MV",              # uit CRV
                "PFW": "PFW",                          # uit PIM K.I. Samen
                "aAa": "aAa",                          # uit PIM K.I. Samen
                "Beta caseïne": "beta caseïne",        # uit PIM K.I. Samen
                "Kappa caseïne": "kappa Caseïne",      # uit PIM K.I. Samen
                "Prijs": "Prijs",                      # uit Prijslijst
                "Prijs gesekst": "Prijs gesekst",      # uit Prijslijst
                "Bt_1": "% betrouwbaarheid",           # uit CRV
                "kgM": "kg melk",                      # uit CRV
                "%V": "% vet",                         # uit CRV
                "%E": "% eiwit",                       # uit CRV
                "kgV": "kg vet",                       # uit CRV
                "kgE": "kg eiwit",                     # uit CRV
                "INET": "INET",                        # uit CRV
                "NVI": "NVI",                          # uit CRV
                "TIP": "TIP",                          # uit Joop Olieman
                "Bt_5": "% betrouwbaar",               # uit CRV (controleer de kolomnaam)
                "F": "frame",                          # uit CRV
                "U": "uier",                           # uit CRV
                "B_6": "benen",                        # uit CRV
                "Ext": "totaal",                       # uit CRV
                "HT": "hoogtemaat",                    # uit CRV
                "VH": "voorhand",                      # uit CRV
                "IH": "inhoud",                        # uit CRV
                "OH": "openheid",                      # uit CRV
                "CS": "conditie score",                # uit CRV
                "KL": "kruisligging",                  # uit CRV
                "KB": "kruisbreedte",                  # uit CRV
                "BA": "beenstand achter",              # uit CRV
                "BZ": "beenstand zij",                 # uit CRV
                "KH": "klauwhoek",                     # uit CRV
                "VB": "voorbeenstand",                 # uit CRV
                "BG": "beengebruik",                   # uit CRV
                "VA": "vooruieraanhechting",            # uit CRV
                "VP": "voorspeenplaatsing",             # uit CRV
                "SL": "speenlengte",                   # uit CRV
                "UD": "uierdiepte",                    # uit CRV
                "AH": "achteruierhoogte",              # uit CRV
                "OB": "ophangband",                    # uit CRV
                "AP": "achterspeenplaatsing",          # uit CRV
                "Geb": "Geboortegemak",                # uit CRV
                "MS": "melksnelheid",                  # uit CRV
                "Cgt": "celgetal",                     # uit CRV
                "Vru": "vruchtbaarheid",               # uit CRV
                "KA": "karakter",                      # uit CRV
                "Ltrh": "laatrijpheid",                # uit CRV
                "Pers": "Persistentie",                # uit CRV
                "Kgh": "klauwgezondheid",              # uit CRV
                "Lvd": "levensduur"                    # uit CRV
            }

            # Bouw de uiteindelijke dataframe op basis van de mapping.
            # Indien een bepaalde bron/kolom niet aanwezig is, wordt de kolom gevuld met lege waarden.
            data = {}
            for kolom_stierenkaart, bron_kolom in kolom_mapping.items():
                if bron_kolom in df_merged.columns:
                    data[kolom_stierenkaart] = df_merged[bron_kolom]
                else:
                    data[kolom_stierenkaart] = None

            df_stierenkaart = pd.DataFrame(data)

            # Converteer de uiteindelijke dataframe naar een Excel-bestand in geheugen
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_stierenkaart.to_excel(writer, sheet_name='stierenkaart', index=False)
            excel_data = output.getvalue()

            st.download_button(
                label="Download gegenereerde stierenkaart",
                data=excel_data,
                file_name="stierenkaart.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Stierenkaart Excel is succesvol gegenereerd!")
