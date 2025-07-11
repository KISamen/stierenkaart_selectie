import os
os.environ["STREAMLIT_WATCH"] = "false"

import streamlit as st
import pandas as pd
import numpy as np
import io

# ----------------------------------------
# 1) Mapping tables
# ----------------------------------------
mapping_table_nl = [
    {"Stierenkaart": "superbevruchter", "Titel in bestand": "Superbevruchter", "Formule": None},
    {"Stierenkaart": "ki-code", "Titel in bestand": "Stiercode NL / KI code", "Formule": None},
    {"Stierenkaart": "naam", "Titel in bestand": "Afkorting stier (zoeknaam)", "Formule": None},
    {"Stierenkaart": "vader", "Titel in bestand": "Roepnaam Vader", "Formule": None},
    {"Stierenkaart": "vaders vader", "Titel in bestand": "Roepnaam Moeders Vader", "Formule": None},
    {"Stierenkaart": "PFW", "Titel in bestand": "PFW code", "Formule": None},
    {"Stierenkaart": "aAa", "Titel in bestand": "AAa code", "Formule": None},
    {"Stierenkaart": "Beta caseine", "Titel in bestand": "Betacasine", "Formule": None},
    {"Stierenkaart": "Kappa caseine", "Titel in bestand": "Kappa-caseine", "Formule": None},
    {"Stierenkaart": "prijs", "Titel in bestand": "Prijs", "Formule": None},
    {"Stierenkaart": "prijs gesekst", "Titel in bestand": "", "Formule": None},
    {"Stierenkaart": "&betrouwbaarheid productie", "Titel in bestand": "Official Production Evaluation in this Country %betrouwbaarheid (Productie-index)", "Formule": None},
    {"Stierenkaart": "kg melk", "Titel in bestand": "Official Production Evalution in this Country KG Melk", "Formule": "/10"},
    {"Stierenkaart": "%vet", "Titel in bestand": "Offical Production Evaluation in this Country %vet", "Formule": "/100"},
    {"Stierenkaart": "%eiwit", "Titel in bestand": "Official Production Evaluation in this County %eiwit", "Formule": "/100"},
    {"Stierenkaart": "kg vet", "Titel in bestand": "Official Production Evaluation in this Country KG vet", "Formule": "/10"},
    {"Stierenkaart": "kg eiwit", "Titel in bestand": "Official Production Evaluation in this Country KG eiwit", "Formule": None},
    {"Stierenkaart": "INET", "Titel in bestand": "Official Production Evaluation in this Country Inet", "Formule": None},
    {"Stierenkaart": "NVI", "Titel in bestand": "Official Production Evaluation in this Country NVI", "Formule": None},
    {"Stierenkaart": "TIP", "Titel in bestand": "", "Formule": None},
    {"Stierenkaart": "%betrouwbaarheid exterieur", "Titel in bestand": "%betrouwbaarheid (exterieur-index)", "Formule": None},
    {"Stierenkaart": "frame", "Titel in bestand": "GENERAL CHARACTERISTICS frame", "Formule": "/100"},
    {"Stierenkaart": "uier", "Titel in bestand": "GENERAL CHARACTERISTICS uier", "Formule": "/100"},
    {"Stierenkaart": "benen", "Titel in bestand": "GENERAL CHARACTERISTICS benen", "Formule": "/100"},
    {"Stierenkaart": "totaal", "Titel in bestand": "GENERAL CHARACTERISTICS totaal (Exterieur-index)", "Formule": "/100"},
    {"Stierenkaart": "hoogtemaat", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY hoogtemaat", "Formule": "/100"},
    {"Stierenkaart": "voorhand", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorhand", "Formule": "/100"},
    {"Stierenkaart": "inhoud", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY inhoud", "Formule": "/100"},
    {"Stierenkaart": "ribvorm", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY openheid", "Formule": "/100"},
    {"Stierenkaart": "conditiescore", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY conditiescore", "Formule": "/100"},
    {"Stierenkaart": "kruisligging", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY kruisligging", "Formule": "/100"},
    {"Stierenkaart": "kruisbreedte", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY kruisbreedte", "Formule": "/100"},
    {"Stierenkaart": "beenstand achter", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achteraanzicht benen", "Formule": "/100"},
    {"Stierenkaart": "beenstand zij", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY beenstand zij", "Formule": "/100"},
    {"Stierenkaart": "klauwhoek", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY klauwhoek", "Formule": "/100"},
    {"Stierenkaart": "voorbeenstand", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorbeenstand", "Formule": "/100"},
    {"Stierenkaart": "beengebruik", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY beengebruik", "Formule": "/100"},
    {"Stierenkaart": "vooruieraanhechting", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY vooruieraanhechting", "Formule": "/100"},
    {"Stierenkaart": "voorspeenplaatsing", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorspeenplaatsing", "Formule": "/100"},
    {"Stierenkaart": "speenlengte", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY speenlengte", "Formule": "/100"},
    {"Stierenkaart": "uierdiepte", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY uierdiepte", "Formule": "/100"},
    {"Stierenkaart": "achteruierhoogte", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achteruierhoogte", "Formule": "/100"},
    {"Stierenkaart": "ophangband", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY ophangband", "Formule": "/100"},
    {"Stierenkaart": "achterspeenplaatsing", "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achterspeenplaatsing", "Formule": "/100"},
    {"Stierenkaart": "geboortegemak", "Titel in bestand": "OFFICIAL CALVING EASE EVALUATION IN THIS COUNTRY geboortegemak", "Formule": "/100"},
    {"Stierenkaart": "melksnelheid", "Titel in bestand": "OFFICIAL MILKING SPEED AND TEMPERAMENT EVALUATION IN THIS COUNTRY melksnelheid", "Formule": "/100"},
    {"Stierenkaart": "celgetal", "Titel in bestand": "OFFICIAL SOMATIC CELL COUNT EVALUATION IN THIS COUNTRY celgetal", "Formule": "/100"},
    {"Stierenkaart": "vruchtbaarheid", "Titel in bestand": "OFFICIAL FEMALE FERTILITY EVALUATION IN THIS COUNTRY vruchtbaarheid", "Formule": "/100"},
    {"Stierenkaart": "karakter", "Titel in bestand": "OFFICIAL MILKING SPEED AND TEMPERAMENT EVALUATION IN THIS COUNTRY karakter", "Formule": "/100"},
    {"Stierenkaart": "laatrijpheid", "Titel in bestand": "OFFICIAL CALVING EASE EVALUATION IN THIS COUNTRY laatrijpheid", "Formule": "/100"},
    {"Stierenkaart": "persistentie", "Titel in bestand": "", "Formule": "/100"},
    {"Stierenkaart": "klauwgezondheid", "Titel in bestand": "OFFICIAL CLAW HEALTH EVALUATION IN THIS COUNTRY klauwgezondheid", "Formule": "/100"},
    {"Stierenkaart": "levensduur", "Titel in bestand": "OFFICIAL CALF LIVABILITY EVALUATION IN THIS COUNTRY levensduur", "Formule": "/100"}
]

mapping_table_canada = [
    {"Stierenkaart": "ki-code", "Titel in bestand": "Stiercode NL / KI code", "Formule": None},
    {"Stierenkaart": "Name",    "Titel in bestand": "Afkorting stier (zoeknaam)", "Formule": None},
    {"Stierenkaart": "Pedigree sire",        "Titel in bestand": "Roepnaam Vader", "Formule": None},
    {"Stierenkaart": "Pedigree mat grandsire","Titel in bestand": "Roepnaam Vaders Vader", "Formule": None},
    {"Stierenkaart": "aAa",     "Titel in bestand": "AAa code", "Formule": None},
    {"Stierenkaart": "prijs",   "Titel in bestand": "Prijs", "Formule": None},
    {"Stierenkaart": "prijs gesekst", "Titel in bestand": "", "Formule": None},
    {"Stierenkaart": "%reliability","Titel in bestand":"Official Production Evaluation in this Country %betrouwbaarheid (Productie-index)","Formule":None},
    {"Stierenkaart": "kg milk",       "Titel in bestand": "Official Production Evalution in this Country KG Melk", "Formule": "/10"},
    {"Stierenkaart": "%fat",          "Titel in bestand": "Offical Production Evaluation in this Country %vet", "Formule": "/100"},
    {"Stierenkaart": "%protein",      "Titel in bestand": "Official Production Evaluation in this County %eiwit", "Formule": "/100"},
    {"Stierenkaart": "kg fat",         "Titel in bestand": "Official Production Evaluation in this Country KG vet", "Formule": "/10"},
    {"Stierenkaart": "kg protein",     "Titel in bestand": "Official Production Evaluation in this Country KG eiwit", "Formule": None},
    {"Stierenkaart": "%reliability conformation traits","Titel in bestand":"%betrouwbaarheid (exterieur-index)","Formule":None},
    {"Stierenkaart": "frame",         "Titel in bestand": "GENERAL CHARACTERISTICS frame", "Formule": "/100"},
    {"Stierenkaart": "udder",         "Titel in bestand": "GENERAL CHARACTERISTICS uier", "Formule": "/100"},
    {"Stierenkaart": "feet & legs",   "Titel in bestand": "GENERAL CHARACTERISTICS benen", "Formule": "/100"},
    {"Stierenkaart": "final score",   "Titel in bestand": "GENERAL CHARACTERISTICS totaal (Exterieur-index)", "Formule": "/100"},
    {"Stierenkaart": "stature",       "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY hoogtemaat", "Formule": "/100"},
    {"Stierenkaart": "chestwidth",    "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorhand", "Formule": "/100"},
    {"Stierenkaart": "body depth",    "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY inhoud", "Formule": "/100"},
    {"Stierenkaart": "angularity",    "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY openheid", "Formule": "/100"},
    {"Stierenkaart": "condition score","Titel in bestand":"OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY conditiescore","Formule":"/100"},
    {"Stierenkaart": "rump angle",    "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY kruisligging", "Formule": "/100"},
    {"Stierenkaart": "rump width",    "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY kruisbreedte", "Formule": "/100"},
    {"Stierenkaart": "rear legs rear view","Titel in bestand":"OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achteraanzicht benen","Formule":"/100"},
    {"Stierenkaart": "rear leg set",  "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY beenstand zij", "Formule": "/100"},
    {"Stierenkaart": "foot angle",    "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY klauwhoek", "Formule": "/100"},
    {"Stierenkaart": "front feet orientation","Titel in bestand":"OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorbeenstand","Formule":"/100"},
    {"Stierenkaart": "locomotion",    "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY beengebruik", "Formule": "/100"},
    {"Stierenkaart": "fore udder attachment","Titel in bestand":"OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY vooruieraanhechting","Formule":"/100"},
    {"Stierenkaart": "fore teat placement","Titel in bestand":"OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY voorspeenplaatsing","Formule":"/100"},
    {"Stierenkaart": "teat length",   "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY speenlengte", "Formule": "/100"},
    {"Stierenkaart": "udder depth",   "Titel in bestand": "OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY uierdiepte", "Formule": "/100"},
    {"Stierenkaart": "rear udder height","Titel in bestand":"OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achteruierhoogte","Formule":"/100"},
    {"Stierenkaart": "central ligament","Titel in bestand":"OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY ophangband","Formule":"/100"},
    {"Stierenkaart": "rear teat placement","Titel in bestand":"OFFICIAL CONFORMATION EVALUATION IN THIS COUNTRY achterspeenplaatsing","Formule":"/100"},
    {"Stierenkaart": "persistency","Titel in bestand":"","Formule":"/100"},
    {"Stierenkaart": "calving ease","Titel in bestand":"OFFICIAL CALVING EASE EVALUATION IN THIS COUNTRY geboortegemak","Formule":"/100"},
    {"Stierenkaart": "milking speed","Titel in bestand":"OFFICIAL MILKING SPEED AND TEMPERAMENT EVALUATION IN THIS COUNTRY melksnelheid","Formule":"/100"},
    {"Stierenkaart": "somatic cell count","Titel in bestand":"OFFICIAL SOMATIC CELL COUNT EVALUATION IN THIS COUNTRY celgetal","Formule":"/100"},
    {"Stierenkaart": "female fertility","Titel in bestand":"OFFICIAL FEMALE FERTILITY EVALUATION IN THIS COUNTRY vruchtbaarheid","Formule":"/100"},
    {"Stierenkaart": "temperament","Titel in bestand":"OFFICIAL MILKING SPEED AND TEMPERAMENT EVALUATION IN THIS COUNTRY karakter","Formule":"/100"},
    {"Stierenkaart": "maturity rate","Titel in bestand":"OFFICIAL CALVING EASE EVALUATION IN THIS COUNTRY laatrijpheid","Formule":"/100"},
    {"Stierenkaart": "hoofhealth","Titel in bestand":"OFFICIAL CLAW HEALTH EVALUATION IN THIS COUNTRY klauwgezondheid","Formule":"/100"},
    {"Stierenkaart": "Beta caseine","Titel in bestand":"Betacasine","Formule":None},
    {"Stierenkaart": "Kappa caseine","Titel in bestand":"Kappa-caseine","Formule":None},
    {"Stierenkaart": "Superbevruchter","Titel in bestand":"Superbevruchter","Formule":None}
]

# ----------------------------------------
# 2) Hulpfuncties
# ----------------------------------------
def load_excel(file):
    try:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Fout bij laden Excel: {e}")
        return None

def custom_sort_ras(df):
    if "Ras" not in df.columns:
        return df
    naam_col = "naam" if "naam" in df.columns else ("Name" if "Name" in df.columns else None)
    order_map = {"Holstein zwartbont": 1, "Red Holstein": 2}
    df["ras_sort"] = df["Ras"].map(order_map).fillna(3)
    df_sorted = df.sort_values(by=["ras_sort", naam_col], ascending=True)
    df_sorted.drop(columns=["ras_sort"], inplace=True)
    return df_sorted

def create_top5_table(df):
    fokwaarden = ["geboortegemak","celgetal","vruchtbaarheid","klauwgezondheid","uier","benen"]
    blocks = []
    if "Ras" not in df.columns:
        return pd.DataFrame()
    df_nl = df[df["Ras"].isin(["Holstein zwartbont","Holstein zwartbont + RF","Red Holstein"])].copy()
    for fok in fokwaarden:
        header = {"Fokwaarde":fok,"zwartbont_stier":"Stier","zwartbont_value":"Waarde","roodbont_stier":"Stier","roodbont_value":"Waarde"}
        blocks.append(header)
        df_z = df_nl[df_nl["Ras"]=="Holstein zwartbont"].copy()
        df_r = df_nl[df_nl["Ras"]=="Red Holstein"].copy()
        for grp, cols in [(df_z,("zwartbont_stier","zwartbont_value")), (df_r,("roodbont_stier","roodbont_value"))]:
            grp[fok]=pd.to_numeric(grp[fok],errors="coerce")
            grp=grp.sort_values(by=fok,ascending=False).head(5)
            for _,row in grp.iterrows():
                blocks.append({"Fokwaarde":"", cols[0]:str(row[naam_col]), cols[1]:str(row[fok]), **{k:"" for k in header if k not in cols}})
        blocks.append({k:"" for k in header})
    return pd.DataFrame(blocks)

# ----------------------------------------
# 3) Main
# ----------------------------------------
def main():
    st.set_page_config(layout="wide")
    st.title("Stierenkaart Generator")

    taal = st.selectbox("Kies stierenkaart type:",["Nederland","Canada"])
    mapping_table = mapping_table_nl if taal=="Nederland" else mapping_table_canada

    uploaded = st.file_uploader("Upload PIM (.xlsx)",type=["xlsx"])
    if not uploaded:
        return
    df = load_excel(uploaded)
    if df is None:
        return

    # 3.1 Mappen & schoonmaken
    data = {}
    for m in mapping_table:
        out=m["Stierenkaart"]; src=m["Titel in bestand"]; f=m["Formule"]
        if src and src in df.columns:
            col=df[src].replace([99999,"+999"],pd.NA)
            if f: 
                col=pd.to_numeric(col,errors="coerce")
                col = col/10 if f=="/10" else col/100
            data[out]=col.astype(str).replace("nan","")
        else:
            data[out]=[""]*len(df)

    df_map=pd.DataFrame(data)

    # 3.2 pinkenstier (NL)
    if taal=="Nederland" and "geboortegemak" in df_map:
        df_map["pinkenstier"]=df_map["geboortegemak"].astype(float).apply(lambda x: "p" if x>1 else "")
    else:
        df_map["pinkenstier"]=""

    # 3.3 kolomnamen
    naam_col = "naam" if taal=="Nederland" else "Name"
    ki_col   = "ki-code"

    # 3.4 volgorde
    volg=[ "superbevruchter", ki_col, naam_col, "pinkenstier"] + [m["Stierenkaart"] for m in mapping_table if m["Stierenkaart"] not in ["superbevruchter",ki_col,naam_col]]
    volg=[c for c in volg if c in df_map]
    overige=[c for c in df_map if c not in volg]
    df_map=df_map[volg+overige]

    # 3.5 Display
    if ki_col in df_map and naam_col in df_map:
        df_map["Display"]=df_map[ki_col]+" - "+df_map[naam_col]
        keuze=st.multiselect("Selecteer stieren", df_map["Display"])
        if keuze:
            codes=[x.split(" - ")[0] for x in keuze]
            sel=df_map[df_map[ki_col].isin(codes)].copy()
            sel=custom_sort_ras(sel)
            st.dataframe(sel,use_container_width=True)
            # top5
            if taal=="Nederland":
                top5=create_top5_table(sel)
                st.subheader("Top 5 per ras")
                st.dataframe(top5,use_container_width=True)
            buf=io.BytesIO()
            with pd.ExcelWriter(buf,engine="openpyxl") as w:
                sel.to_excel(w,sheet_name="Stierenkaart",index=False)
                if taal=="Nederland":
                    top5.to_excel(w,sheet_name="Top5",index=False)
            st.download_button("Download Excel",data=buf.getvalue(),file_name="stierenkaart.xlsx")
        else:
            st.info("Selecteer stieren om te tonen.")
    else:
        st.warning(f"Kolommen {ki_col!r} en/of {naam_col!r} ontbreken.")

if __name__=="__main__":
    main()
