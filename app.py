import streamlit as st
import pandas as pd
import io

# -------------------------------------------
# Functie om het PIM-bestand te laden
# -------------------------------------------
def load_excel(file):
    try:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Fout bij laden Excel: {e}")
        return None

# -------------------------------------------
# Sorteer stieren per ras
# -------------------------------------------
def custom_sort_ras(df):
    if "Ras" not in df.columns:
        df["Ras"] = ""
    if "Stier" not in df.columns:
        df["Stier"] = ""
    order_map = {"Holstein zwartbont": 1, "Red Holstein": 2}
    df["ras_sort"] = df["Ras"].map(order_map).fillna(3)
    df_sorted = df.sort_values(by=["ras_sort", "Stier"], ascending=True)
    df_sorted.drop(columns=["ras_sort"], inplace=True)
    return df_sorted

# -------------------------------------------
# Bouw Top 5-tabellen
# -------------------------------------------
def create_top5_table(df):
    fokwaarden = [
        "Geboortegemak",
        "celgetal",
        "vruchtbaarheid",
        "klauwgezondheid",
        "uier",
        "benen"
    ]

    blocks = []
    if df.empty:
        return pd.DataFrame()

    df["Ras_clean"] = df["Ras"].astype(str).str.strip().str.lower()
    df = df[df["Ras_clean"].isin([
        "holstein zwartbont",
        "holstein zwartbont + rf",
        "red holstein"
    ])].copy()

    for fok in fokwaarden:
        if fok not in df.columns:
            df[fok] = pd.NA

        block = []
        header_row = {
            "Fokwaarde": fok,
            "zwartbont_stier": "Stier",
            "zwartbont_value": "Waarde",
            "roodbont_stier": "Stier",
            "roodbont_value": "Waarde"
        }
        block.append(header_row)

        df_z = df[df["Ras_clean"].isin(["holstein zwartbont", "holstein zwartbont + rf"])].copy()
        df_z[fok] = pd.to_numeric(df_z[fok], errors='coerce')
        df_z = df_z.sort_values(by=fok, ascending=False)

        df_r = df[df["Ras_clean"].str.contains("red holstein")].copy()
        df_r[fok] = pd.to_numeric(df_r[fok], errors='coerce')
        df_r = df_r.sort_values(by=fok, ascending=False)

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

        block.append({
            "Fokwaarde": "",
            "zwartbont_stier": "",
            "zwartbont_value": "",
            "roodbont_stier": "",
            "roodbont_value": ""
        })

        blocks.extend(block)

    return pd.DataFrame(blocks)

# -------------------------------------------
# Streamlit app
# -------------------------------------------
def main():
    st.set_page_config(layout="wide")
    st.title("Stierenkaart Generator (vereenvoudigd)")

    uploaded_file = st.file_uploader("Upload PIM K.I. Samen.xlsx", type=["xlsx"])

    if uploaded_file:
        df = load_excel(uploaded_file)

        if df is not None:
            st.success(f"Bestand ingelezen met {len(df)} rijen en {len(df.columns)} kolommen.")

            if st.checkbox("Toon kolomnamen"):
                st.write(df.columns.tolist())

            if "KI-Code" in df.columns and "Stier" in df.columns:
                df["KI-Code"] = df["KI-Code"].astype(str).str.strip().str.upper()
                df["Display"] = df["KI-Code"] + " - " + df["Stier"].astype(str)

                selected_display = st.multiselect(
                    "Selecteer stieren:",
                    options=df["Display"].tolist()
                )

                if selected_display:
                    selected_codes = [x.split(" - ")[0] for x in selected_display]
                    df_selected = df[df["KI-Code"].isin(selected_codes)].copy()
                    df_selected = custom_sort_ras(df_selected)

                    st.subheader("Geselecteerde stieren")
                    st.dataframe(df_selected, use_container_width=True)

                    df_top5 = create_top5_table(df_selected)
                    if not df_top5.empty:
                        st.subheader("Top 5-tabellen per fokwaarde")
                        st.dataframe(df_top5, use_container_width=True)

                    # Excel download
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_selected.to_excel(writer, sheet_name='Stierenkaart', index=False)
                        if not df_top5.empty:
                            df_top5.to_excel(writer, sheet_name='Top5_per_ras', index=False)

                    st.download_button(
                        label="Download selectie + Top 5-tabellen",
                        data=output.getvalue(),
                        file_name="stierenkaart_selectie.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info("Selecteer één of meer stieren om de gegevens te zien en te downloaden.")
            else:
                st.warning("Kolommen 'KI-Code' en/of 'Stier' niet gevonden in het PIM-bestand.")
        else:
            st.error("Kon het bestand niet inlezen.")
    else:
        st.info("Upload eerst het PIM-bestand.")

if __name__ == "__main__":
    main()

