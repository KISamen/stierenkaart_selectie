import streamlit as st
import pandas as pd
import io

# --- Mapping tables per taal ---
# NL-stierenkaart (jouw originele mapping table)
mapping_table_nl = [
    {"Titel in bestand": "KI_Code", "Stierenkaart": "KI-code", "Waar te vinden": ""},
    {"Titel in bestand": "Eigenaarscode", "Stierenkaart": "Eigenaarscode", "Waar te vinden": ""},
    {"Titel in bestand": "Stiernummer", "Stierenkaart": "Stiernummer", "Waar te vinden": ""},
    {"Titel in bestand": "Stiernaam", "Stierenkaart": "Stier", "Waar te vinden": ""},
    {"Titel in bestand": "Erf-fact", "Stierenkaart": "Erf-fact", "Waar te vinden": ""},
    {"Titel in bestand": "Vader", "Stierenkaart": "Afstamming V", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "M-vader", "Stierenkaart": "Afstamming MV", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "PFW", "Stierenkaart": "PFW", "Waar te vinden": "PIM K.I. SAMEN"},
    {"Titel in bestand": "AAa code", "Stierenkaart": "aAa", "Waar te vinden": "PIM K.I. SAMEN"},
    {"Titel in bestand": "Betacasine", "Stierenkaart": "beta caseïne", "Waar te vinden": "PIM K.I. SAMEN"},
    {"Titel in bestand": "Kappa-caseine", "Stierenkaart": "kappa Caseïne", "Waar te vinden": "PIM K.I. SAMEN"},
    {"Titel in bestand": "Prijs", "Stierenkaart": "Prijs", "Waar te vinden": "Prijslijst"},
    {"Titel in bestand": "Prijs gesekst", "Stierenkaart": "Prijs gesekst", "Waar te vinden": "Prijslijst"},
    {"Titel in bestand": "Bt_1", "Stierenkaart": "% betrouwbaarheid", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "kgM", "Stierenkaart": "kg melk", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "%V", "Stierenkaart": "% vet", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "%E", "Stierenkaart": "% eiwit", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "kgV", "Stierenkaart": "kg vet", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "kgE", "Stierenkaart": "kg eiwit", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "INET", "Stierenkaart": "INET", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "NVI", "Stierenkaart": "NVI", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "TIP", "Stierenkaart": "TIP", "Waar te vinden": "Bronbestand Joop Olieman"},
    {"Titel in bestand": "Bt_5", "Stierenkaart": "% betrouwbaar", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "F", "Stierenkaart": "frame", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "U", "Stierenkaart": "uier", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "B_6", "Stierenkaart": "benen", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Ext", "Stierenkaart": "totaal", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "HT", "Stierenkaart": "hoogtemaat", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "VH", "Stierenkaart": "voorhand", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "IH", "Stierenkaart": "inhoud", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "OH", "Stierenkaart": "openheid", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "CS", "Stierenkaart": "conditie score", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "KL", "Stierenkaart": "kruisligging", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "KB", "Stierenkaart": "kruisbreedte", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "BA", "Stierenkaart": "beenstand achter", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "BZ", "Stierenkaart": "beenstand zij", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "KH", "Stierenkaart": "klauwhoek", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "VB", "Stierenkaart": "voorbeenstand", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "BG", "Stierenkaart": "beengebruik", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "VA", "Stierenkaart": "vooruieraanhechting", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "VP", "Stierenkaart": "voorspeenplaatsing", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "SL", "Stierenkaart": "speenlengte", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "UD", "Stierenkaart": "uierdiepte", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "AH", "Stierenkaart": "achteruierhoogte", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "OB", "Stierenkaart": "ophangband", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "AP", "Stierenkaart": "achterspeenplaatsing", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Rasomschrijving", "Stierenkaart": "Ras", "Waar te vinden": "PIM K.I. SAMEN"},
    {"Titel in bestand": "Geb", "Stierenkaart": "Geboortegemak", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "MS", "Stierenkaart": "melksnelheid", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Cgt", "Stierenkaart": "celgetal", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Vru", "Stierenkaart": "vruchtbaarheid", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "KA", "Stierenkaart": "karakter", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Ltrh", "Stierenkaart": "laatrijpheid", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Pers", "Stierenkaart": "Persistentie", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Kgh", "Stierenkaart": "klauwgezondheid", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Lvd", "Stierenkaart": "levensduur", "Waar te vinden": "Bronbestand CRV"}
]

# --- Mapping tables per taal ---
# NL-stierenkaart (jouw originele mapping table)
mapping_table_vlaams = [
    {"Titel in bestand": "KI_Code", "Stierenkaart": "KI-code", "Waar te vinden": ""},
    {"Titel in bestand": "Eigenaarscode", "Stierenkaart": "Eigenaarscode", "Waar te vinden": ""},
    {"Titel in bestand": "Stiernummer", "Stierenkaart": "Stiernummer", "Waar te vinden": ""},
    {"Titel in bestand": "Stiernaam", "Stierenkaart": "Stier", "Waar te vinden": ""},
    {"Titel in bestand": "Erf-fact", "Stierenkaart": "Erf-fact", "Waar te vinden": ""},
    {"Titel in bestand": "Vader", "Stierenkaart": "Afstamming V", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "M-vader", "Stierenkaart": "Afstamming MV", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "PFW", "Stierenkaart": "PFW", "Waar te vinden": "PIM K.I. SAMEN"},
    {"Titel in bestand": "AAa code", "Stierenkaart": "aAa", "Waar te vinden": "PIM K.I. SAMEN"},
    {"Titel in bestand": "Betacasine", "Stierenkaart": "beta caseïne", "Waar te vinden": "PIM K.I. SAMEN"},
    {"Titel in bestand": "Kappa-caseine", "Stierenkaart": "kappa Caseïne", "Waar te vinden": "PIM K.I. SAMEN"},
    {"Titel in bestand": "Prijs", "Stierenkaart": "Prijs", "Waar te vinden": "Prijslijst"},
    {"Titel in bestand": "Prijs gesekst", "Stierenkaart": "Prijs gesekst", "Waar te vinden": "Prijslijst"},
    {"Titel in bestand": "Bt_1", "Stierenkaart": "% betrouwbaarheid", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "kgM", "Stierenkaart": "kg melk", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "%V", "Stierenkaart": "% vet", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "%E", "Stierenkaart": "% eiwit", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "kgV", "Stierenkaart": "kg vet", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "kgE", "Stierenkaart": "kg eiwit", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "INET", "Stierenkaart": "INET", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "NVI", "Stierenkaart": "NVI", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "TIP", "Stierenkaart": "TIP", "Waar te vinden": "Bronbestand Joop Olieman"},
    {"Titel in bestand": "Bt_5", "Stierenkaart": "% betrouwbaar", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "F", "Stierenkaart": "frame", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "U", "Stierenkaart": "uier", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "B_6", "Stierenkaart": "benen", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Ext", "Stierenkaart": "totaal", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "HT", "Stierenkaart": "hoogtemaat", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "VH", "Stierenkaart": "voorhand", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "IH", "Stierenkaart": "inhoud", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "OH", "Stierenkaart": "openheid", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "CS", "Stierenkaart": "conditie score", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "KL", "Stierenkaart": "kruisligging", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "KB", "Stierenkaart": "kruisbreedte", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "BA", "Stierenkaart": "beenstand achter", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "BZ", "Stierenkaart": "beenstand zij", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "KH", "Stierenkaart": "klauwhoek", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "VB", "Stierenkaart": "voorbeenstand", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "BG", "Stierenkaart": "beengebruik", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "VA", "Stierenkaart": "vooruieraanhechting", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "VP", "Stierenkaart": "voorspeenplaatsing", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "SL", "Stierenkaart": "speenlengte", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "UD", "Stierenkaart": "uierdiepte", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "AH", "Stierenkaart": "achteruierhoogte", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "OB", "Stierenkaart": "ophangband", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "AP", "Stierenkaart": "achterspeenplaatsing", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Rasomschrijving", "Stierenkaart": "Ras", "Waar te vinden": "PIM K.I. SAMEN"},
    {"Titel in bestand": "Geb", "Stierenkaart": "Geboortegemak", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "MS", "Stierenkaart": "melksnelheid", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Cgt", "Stierenkaart": "celgetal", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Vru", "Stierenkaart": "vruchtbaarheid", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "KA", "Stierenkaart": "karakter", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Ltrh", "Stierenkaart": "laatrijpheid", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Pers", "Stierenkaart": "Persistentie", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Kgh", "Stierenkaart": "klauwgezondheid", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Lvd", "Stierenkaart": "levensduur", "Waar te vinden": "Bronbestand CRV"}
]


mapping_table_waals = [
    {"Titel in bestand": "KI_Code", "Stierenkaart": "code centre", "Waar te vinden": ""},
    {"Titel in bestand": "Eigenaarscode", "Stierenkaart": "Eigenaarscode", "Waar te vinden": ""},
    {"Titel in bestand": "Stiernummer", "Stierenkaart": "Stiernummer", "Waar te vinden": ""},
    {"Titel in bestand": "Stiernaam", "Stierenkaart": "taureau", "Waar te vinden": ""},
    {"Titel in bestand": "Vader", "Stierenkaart": "pedigree", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "M-vader", "Stierenkaart": "pedigree", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "PFW", "Stierenkaart": "PFW", "Waar te vinden": "PIM K.I. SAMEN"},
    {"Titel in bestand": "AAa code", "Stierenkaart": "aAa", "Waar te vinden": "PIM K.I. SAMEN"},
    {"Titel in bestand": "Betacasine", "Stierenkaart": "beta caseïne", "Waar te vinden": "PIM K.I. SAMEN"},
    {"Titel in bestand": "Kappa-caseine", "Stierenkaart": "kappa Caseïne", "Waar te vinden": "PIM K.I. SAMEN"},
    {"Titel in bestand": "Prijs", "Stierenkaart": "Tarif", "Waar te vinden": "Prijslijst"},
    {"Titel in bestand": "Prijs gesekst", "Stierenkaart": "Tarif Sexée", "Waar te vinden": "Prijslijst"},
    {"Titel in bestand": "Bt_1", "Stierenkaart": "%cd", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "kgM", "Stierenkaart": "kg lait", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "%V", "Stierenkaart": "%TB", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "%E", "Stierenkaart": "%TA", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "kgV", "Stierenkaart": "kgMG", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "kgE", "Stierenkaart": "kgMA", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "INET", "Stierenkaart": "INET", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "NVI", "Stierenkaart": "NVI", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "TIP", "Stierenkaart": "TIP", "Waar te vinden": "Bronbestand Joop Olieman"},
    {"Titel in bestand": "Bt_5", "Stierenkaart": "%cd", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "F", "Stierenkaart": "frame / charpente", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "U", "Stierenkaart": "pis", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "B_6", "Stierenkaart": "membres", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Ext", "Stierenkaart": "total type", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "HT", "Stierenkaart": "hauteur sacrum", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "VH", "Stierenkaart": "prof. poitrine", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "IH", "Stierenkaart": "capacité", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "OH", "Stierenkaart": "caractere laitier", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "CS", "Stierenkaart": "condition", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "KL", "Stierenkaart": "inclinaison bassin", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "KB", "Stierenkaart": "largeur bassin", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "BA", "Stierenkaart": "assiette membres arr.", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "BZ", "Stierenkaart": "angle jarret", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "KH", "Stierenkaart": "klauwhoek", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "VB", "Stierenkaart": "voorbeenstand", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "BG", "Stierenkaart": "beengebruik", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "VA", "Stierenkaart": "vooruieraanhechting", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "VP", "Stierenkaart": "attache avant", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "SL", "Stierenkaart": "longueur trayonas av.", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "UD", "Stierenkaart": "profondeur", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "AH", "Stierenkaart": "attache arriere", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "OB", "Stierenkaart": "ligament", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "AP", "Stierenkaart": "placement trayons arriére", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Rasomschrijving", "Stierenkaart": "Ras", "Waar te vinden": "PIM K.I. SAMEN"},
    {"Titel in bestand": "Geb", "Stierenkaart": "facilité de vêlage", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "MS", "Stierenkaart": "vitesse de traite", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Cgt", "Stierenkaart": "cellulles*", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Vru", "Stierenkaart": "fertilité", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "KA", "Stierenkaart": "caractere", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Ltrh", "Stierenkaart": "tardivité", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Pers", "Stierenkaart": "persistance", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Kgh", "Stierenkaart": "santé des pieds", "Waar te vinden": "Bronbestand CRV"},
    {"Titel in bestand": "Lvd", "Stierenkaart": "longévité", "Waar te vinden": "Bronbestand CRV"}
]

# Voorlopige placeholders voor andere talen (pas de titels aan indien nodig)     
mapping_table_engels  = mapping_table_nl.copy()    
mapping_table_duits   = mapping_table_nl.copy()    
mapping_table_canadese = mapping_table_nl.copy()  

# Verzamel alle mapping tables in een dictionary
mapping_tables = {
    "NL": mapping_table_nl,
    "Vlaams": mapping_table_vlaams,
    "Waals": mapping_table_waals,
    "Engels": mapping_table_engels,
    "Duits": mapping_table_duits,
    "Canadese": mapping_table_canadese
}

# --- Einde Mapping tables ---

def load_excel(file):
    try:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Fout bij laden bestand: {e}")
        return None

def custom_sort_ras(df):
    # Zorg dat de kolommen 'Ras' en 'Stier' aanwezig zijn.
    if "Ras" not in df.columns:
        df["Ras"] = ""
    if "Stier" not in df.columns:
        df["Stier"] = ""
    order_map = {"Holstein zwartbont": 1, "Red Holstein": 2}
    df["ras_sort"] = df["Ras"].map(order_map).fillna(3)
    df_sorted = df.sort_values(by=["ras_sort", "Stier"], ascending=True)
    df_sorted.drop(columns=["ras_sort"], inplace=True)
    df_with_header = pd.DataFrame()
    first_group = True
    for ras, group in df_sorted.groupby("Ras"):
        if not first_group:
            df_with_header = pd.concat([df_with_header, pd.DataFrame([{}])], ignore_index=True)
        header_row = {col: "" for col in df_sorted.columns}
        header_row["Ras"] = ras
        df_with_header = pd.concat([df_with_header, pd.DataFrame([header_row])], ignore_index=True)
        df_with_header = pd.concat([df_with_header, group], ignore_index=True)
        first_group = False
    return df_with_header

def create_top5_table(df):
    # Maak een extra kolom met een 'clean' versie van Ras (lowercase en getrimd)
    df["Ras_clean"] = df["Ras"].astype(str).str.strip().str.lower()
    fokwaarden = ["Geboortegemak", "celgetal", "vruchtbaarheid", "klauwgezondheid", "uier", "benen"]
    blocks = []
    # Beperk tot stieren met Ras "Holstein zwartbont", "Holstein zwartbont + RF" of "Red Holstein"
    df = df[df["Ras"].isin(["Holstein zwartbont", "Holstein zwartbont + RF", "Red Holstein"])].copy()
    
    for fok in fokwaarden:
        if fok not in df.columns:
            df[fok] = pd.NA
        block = []
        # Header-rij voor deze fokwaarde met aparte kolommen voor stier en waarde
        header_row = {
            "Fokwaarde": fok,
            "zwartbont_stier": "Stier",
            "zwartbont_value": "Waarde",
            "roodbont_stier": "Stier",
            "roodbont_value": "Waarde"
        }
        block.append(header_row)
        
        # Filter voor zwarte stieren: rijen waarbij Ras_clean gelijk is aan "holstein zwartbont" of "holstein zwartbont + rf"
        df_z = df[df["Ras_clean"].isin(["holstein zwartbont", "holstein zwartbont + rf"])].copy()
        df_z[fok] = pd.to_numeric(df_z[fok], errors='coerce')
        df_z = df_z.sort_values(by=fok, ascending=False)
        
        # Filter voor rode stieren: rijen waarbij Ras_clean "red holstein" bevat.
        df_r = df[df["Ras_clean"].str.contains("red holstein")].copy()
        df_r[fok] = pd.to_numeric(df_r[fok], errors='coerce')
        df_r = df_r.sort_values(by=fok, ascending=False)
        
        # Voeg 5 rijen toe met de top 5 waarden per groep
        for i in range(5):
            row = {
                "Fokwaarde": "",
                "zwartbont_stier": "",
                "zwartbont_value": "",
                "roodbont_stier": "",
                "roodbont_value": ""
            }
            if i < len(df_z):
                row["zwartbont_stier"] = str(df_z.iloc[i]["Stier"])
                row["zwartbont_value"] = str(df_z.iloc[i][fok])
            if i < len(df_r):
                row["roodbont_stier"] = str(df_r.iloc[i]["Stier"])
                row["roodbont_value"] = str(df_r.iloc[i][fok])
            block.append(row)
        # Voeg een lege rij toe als scheiding
        block.append({
            "Fokwaarde": "",
            "zwartbont_stier": "",
            "zwartbont_value": "",
            "roodbont_stier": "",
            "roodbont_value": ""
        })
        blocks.extend(block)
    return pd.DataFrame(blocks)

# --- Einde Functies ---

def main():
    st.set_page_config(layout="wide")
    st.title("Stierenkaart Generator")
    st.markdown("""
    **Upload de volgende bestanden:**
    - **Bronbestand CRV DEC2024.xlsx**
    - **PIM K.I. Samen.xlsx**
    - **Prijslijst.xlsx**
    - **Bronbestand Joop Olieman.xlsx**
    
    **Bulk-selectie:**  
    Excelbestand met KI-code in kolom A.
    """)
    
    # Laat de gebruiker kiezen welke stierenkaart (en mapping table) gebruikt moet worden.
    taal_keuze = st.selectbox("Kies stierenkaart taal", options=["NL", "Vlaams", "Waals", "Engels", "Duits", "Canadese"])
    selected_mapping_table = mapping_tables.get(taal_keuze, mapping_table_nl)
    
    # Uploaders voor de vier bestanden
    uploaded_crv = st.file_uploader("Upload Bronbestand CRV DEC2024.xlsx", type=["xlsx"], key="crv")
    uploaded_pim = st.file_uploader("Upload PIM K.I. Samen.xlsx", type=["xlsx"], key="pim")
    uploaded_prijslijst = st.file_uploader("Upload Prijslijst.xlsx", type=["xlsx"], key="prijslijst")
    uploaded_joop = st.file_uploader("Upload Bronbestand Joop Olieman.xlsx", type=["xlsx"], key="joop")
    debug_mode = st.checkbox("Activeer debug", value=False)
    
    # Bulk-selectie: upload en bewaar in session state
    bulk_file = st.file_uploader("Upload bulk selectie bestand (KI-code kolom A)", type=["xlsx"], key="bulk")
    if bulk_file is not None:
        df_bulk = pd.read_excel(bulk_file)
        bulk_codes = df_bulk.iloc[:, 0].astype(str).str.upper().str.strip().tolist()
        st.session_state.bulk_selected_codes = bulk_codes
    else:
        if "bulk_selected_codes" not in st.session_state:
            st.session_state.bulk_selected_codes = []
    
    if "df_stierenkaart" not in st.session_state:
        st.session_state.df_stierenkaart = None
    
    if st.button("Genereer Stierenkaart"):
        if not (uploaded_crv and uploaded_pim and uploaded_prijslijst and uploaded_joop):
            st.error("Upload alle bestanden!")
        else:
            df_crv = load_excel(uploaded_crv)
            df_pim = load_excel(uploaded_pim)
            df_prijslijst = load_excel(uploaded_prijslijst)
            df_joop = load_excel(uploaded_joop)
            
            if any(df is None for df in [df_crv, df_pim, df_prijslijst, df_joop]):
                st.error("Fout bij laden van één of meer bestanden.")
            else:
                # Normaliseer KI-codes
                df_crv["KI_Code"] = df_crv["KI-Code"].astype(str).str.upper().str.strip()
                df_pim["KI_Code"] = df_pim["Stiercode NL / KI code"].astype(str).str.upper().str.strip()
                df_prijslijst["KI_Code"] = df_prijslijst["Artikelnr."].astype(str).str.upper().str.strip()
                df_joop["KI_Code"] = df_joop["Kicode"].astype(str).str.upper().str.strip()
                
                # In PIM: hernoem "PFW code" naar "PFW"
                pfw_col = None
                for col in df_pim.columns:
                    if col.lower() == "pfw code":
                        pfw_col = col
                        break
                if pfw_col:
                    df_pim.rename(columns={pfw_col: "PFW"}, inplace=True)
                    df_pim["PFW"] = df_pim["PFW"].astype(str).str.strip()
                else:
                    st.warning("Kolom 'PFW code' niet gevonden in het PIM-bestand.")
                
                # In Joop: hernoem TIP-kolom naar "TIP"
                tip_col = None
                for col in df_joop.columns:
                    if col.strip().upper() == "TIP":
                        tip_col = col
                        break
                if not tip_col:
                    for col in df_joop.columns:
                        if col.strip().upper().startswith("TIP"):
                            tip_col = col
                            break
                if tip_col:
                    if tip_col != "TIP":
                        df_joop.rename(columns={tip_col: "TIP"}, inplace=True)
                    df_joop["TIP"] = df_joop["TIP"].astype(str).str.strip()
                else:
                    st.warning("Kolom 'TIP' niet gevonden in het Joop-bestand.")
                
                # Voeg tijdelijke key toe
                for df_temp in [df_crv, df_pim, df_prijslijst, df_joop]:
                    df_temp["temp_key"] = df_temp["KI_Code"]
                # Merge de dataframes
                df_merged = pd.merge(df_crv, df_pim, on="temp_key", how="left", suffixes=("", "_pim"))
                df_merged = pd.merge(df_merged, df_prijslijst, on="temp_key", how="left", suffixes=("", "_prijslijst"))
                df_merged = pd.merge(df_merged, df_joop, on="temp_key", how="left", suffixes=("", "_joop"))
                
                if "KI_Code" in df_crv.columns:
                    df_merged["KI_Code"] = df_crv["KI_Code"]
                else:
                    st.error("Kolom 'KI_Code' ontbreekt in het CRV-bestand.")
                
                if debug_mode:
                    st.write("Debug: Kolommen in df_merged:", df_merged.columns.tolist())
                    st.write("Debug: Voorbeeld data PFW:", df_merged[["KI_Code", "PFW"]].head())
                    st.write("Debug: Voorbeeld data TIP:", df_merged[["KI_Code", "TIP"]].head())
                
                # Gebruik de geselecteerde mapping table
                mapping_table = selected_mapping_table
                
                final_data = {}
                for mapping in mapping_table:
                    titel = mapping.get("Titel in bestand", "")
                    std_naam = mapping.get("Stierenkaart", "")
                    if titel in df_merged.columns:
                        final_data[std_naam] = df_merged[titel]
                    else:
                        final_data[std_naam] = ""
                
                df_stierenkaart = pd.DataFrame(final_data)
                df_stierenkaart.fillna("", inplace=True)
                st.session_state.df_stierenkaart = df_stierenkaart
            
            # Handmatige selectie van stieren
            if st.session_state.get("df_stierenkaart") is not None:
                df_stierenkaart = st.session_state.df_stierenkaart
                df_stierenkaart["Display"] = df_stierenkaart["KI-code"] + " - " + df_stierenkaart["Stier"]
                mapping_dict = dict(zip(df_stierenkaart["KI-code"], df_stierenkaart["Display"]))
                
                basic_selection = st.multiselect("Selecteer stieren", options=list(mapping_dict.values()))
                
                with st.expander("Selecteer per ras"):
                    grouped_options = {}
                    for _, row in df_stierenkaart.iterrows():
                        breed = row["Ras"]
                        display = row["Display"]
                        if breed not in grouped_options:
                            grouped_options[breed] = []
                        grouped_options[breed].append(display)
                    for breed in grouped_options:
                        grouped_options[breed] = sorted(list(set(grouped_options[breed])))
                    order_map = {"Holstein zwartbont": 1, "Red Holstein": 2}
                    sorted_breeds = sorted(grouped_options.keys(), key=lambda x: order_map.get(x, 3))
                    
                    st.markdown("### Selectie per ras")
                    per_ras_selection = []
                    for breed in sorted_breeds:
                        st.markdown(f"**{breed}**")
                        sel = st.multiselect(f"Selecteer stieren ({breed})", options=grouped_options[breed], key=f"ms_{breed}")
                        per_ras_selection.extend(sel)
                
                manual_selected_codes = [item.split(" - ")[0] for item in basic_selection + per_ras_selection]
                bulk_selected_codes = st.session_state.bulk_selected_codes if "bulk_selected_codes" in st.session_state else []
                combined_codes = sorted(set(bulk_selected_codes + manual_selected_codes))
                valid_final_display = [mapping_dict.get(code) for code in combined_codes if mapping_dict.get(code)]
                
                final_combined_display = st.multiselect("Gecombineerde selectie:", options=list(mapping_dict.values()), default=valid_final_display)
                final_selected_codes = [item.split(" - ")[0] for item in final_combined_display]
                
                df_selected_filtered = df_stierenkaart[df_stierenkaart["KI-code"].isin(final_selected_codes)]
                df_overig = df_stierenkaart[~df_stierenkaart["KI-code"].isin(final_selected_codes)]
                
                df_selected = custom_sort_ras(df_selected_filtered)
                df_overig = custom_sort_ras(df_overig)
                
                df_top5 = create_top5_table(df_selected_filtered)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_selected.to_excel(writer, sheet_name='Stierenkaart', index=False)
                    df_overig.to_excel(writer, sheet_name='Overige stieren', index=False)
                    df_top5.to_excel(writer, sheet_name='Top 5 per ras', index=False)
                
                st.download_button(
                    label="Download stierenkaart Excel",
                    data=output.getvalue(),
                    file_name="stierenkaart.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == '__main__':
    main()
