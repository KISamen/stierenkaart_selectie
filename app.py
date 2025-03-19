import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide")
st.title("🐄 Stierenkaart Generator")

# Stap 1: Upload-bestanden interface
uploaded_crv = st.file_uploader("Upload Bronbestand CRV DEC2024.xlsx", type=["xlsx"])
uploaded_pim = st.file_uploader("Upload PIM K.I. Samen.xlsx", type=["xlsx"])
uploaded_prijslijst = st.file_uploader("Upload Prijslijst.xlsx", type=["xlsx"])
uploaded_joop = st.file_uploader("Upload Bronbestand Joop Olieman.xlsx", type=["xlsx"])

debug_mode = st.checkbox("Activeer debug", value=False)

def load_excel(file):
    try:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Fout bij het laden: {e}")
        return None

# Stap 2: Data genereren uit uploads
if st.button("📥 Genereer Stierenkaart"):
    if not (uploaded_crv and uploaded_pim and uploaded_prijslijst and uploaded_joop):
        st.error("Upload alle bestanden!")
    else:
        df_crv = load_excel(uploaded_crv)
        df_pim = load_excel(uploaded_pim)
        df_prijslijst = load_excel(uploaded_prijslijst)
        df_joop = load_excel(uploaded_joop)

        if any(df is None for df in [df_crv, df_pim, df_prijslijst, df_joop]):
            st.error("Fout bij laden bestanden.")
        else:
            df_crv["KI_Code"] = df_crv["KI-Code"].astype(str).str.upper().str.strip()
            df_pim["KI_Code"] = df_pim["Stiercode NL / KI code"].astype(str).str.upper().str.strip()
            df_prijslijst["KI_Code"] = df_prijslijst["Artikelnr."].astype(str).str.upper().str.strip()
            df_joop["KI_Code"] = df_joop["Kicode"].astype(str).str.upper().str.strip()

            for df in [df_crv, df_pim, df_prijslijst, df_joop]:
                df["temp_key"] = df["KI_Code"]

            df_merged = pd.merge(df_crv, df_pim, on="temp_key", how="left", suffixes=("", "_pim"))
            df_merged = pd.merge(df_merged, df_prijslijst, on="temp_key", how="left")
            df_merged = pd.merge(df_merged, df_joop, on="temp_key", how="left")
            df_merged["KI_Code"] = df_crv["KI_Code"]
            df_merged.drop(columns=["temp_key"], inplace=True)

            final_cols = ["KI_Code", "Stiernaam", "Vader", "M-vader", "PFW", "AAa code",
                          "Betacasine", "Kappa-caseine", "Prijs", "Prijs gesekst", "NVI", "TIP"]
            df_stierenkaart = df_merged[final_cols].copy()
            df_stierenkaart.rename(columns={"KI_Code": "KI-code", "Stiernaam": "Stier"}, inplace=True)
            df_stierenkaart["Stier"] = df_stierenkaart["Stier"].str.upper()

            df_stierenkaart.fillna("", inplace=True)

            st.session_state.df_stierenkaart = df_stierenkaart

# Stap 3: Selectie-interface (bulk + handmatig)
if "df_stierenkaart" in st.session_state:
    df_stierenkaart = st.session_state.df_stierenkaart

    df_stierenkaart["Display"] = df_stierenkaart["KI-code"] + " - " + df_stierenkaart["Stier"]
    options = sorted(df_stierenkaart["Display"].dropna().unique().tolist())

    bulk_file = st.file_uploader("Upload bulk-selectiebestand (kolom A: KI-code)", type=["xlsx"], key="bulk")

    bulk_selected_codes = []
    if bulk_file:
        df_bulk = pd.read_excel(bulk_file)
        bulk_selected_codes = df_bulk.iloc[:,0].astype(str).str.upper().str.strip().tolist()
        st.session_state.bulk_selected = bulk_selected_codes
        if debug_mode:
            st.write("Bulk geselecteerd:", bulk_selected_codes)

    else:
        bulk_selected_codes = st.session_state.get("bulk_selected", [])

    manual_selected_display = st.multiselect("Voeg handmatig stieren toe:", options=options)
    manual_selected_codes = [item.split(" - ")[0] for item in manual_selected_display]

    combined_codes = sorted(set(bulk_selected_codes).union(set(manual_selected_codes)))
    mapping_dict = dict(zip(df_stierenkaart["KI-code"], df_stierenkaart["Display"]))
    final_display = [mapping_dict.get(code) for code in combined_codes if mapping_dict.get(code)]

    valid_final_display = [item for item in final_display if item in options]

    final_combined_display = st.multiselect(
        "Gecombineerde selectie (bulk + handmatig):",
        options=options,
        default=valid_final_display
    )

    final_selected_codes = [item.split(" - ")[0] for item in final_combined_display]

    if debug_mode:
        st.write("Final selectie:", final_selected_codes)

    df_selected = df_stierenkaart[df_stierenkaart["KI-code"].isin(final_selected_codes)]
    df_overig = df_stierenkaart[~df_stierenkaart["KI-code"].isin(final_selected_codes)]

    # Stap 4: Excel-export genereren
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_selected.to_excel(writer, sheet_name='Stierenkaart', index=False)
        df_overig.to_excel(writer, sheet_name='Overige stieren', index=False)

    excel_data = output.getvalue()
    st.download_button(
        label="💾 Download stierenkaart Excel",
        data=excel_data,
        file_name="stierenkaart.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("✅ Excel-bestand succesvol gegenereerd!")

