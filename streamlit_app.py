import pandas as pd
import streamlit as st
import io

def process_aly_sheet(df, objekt_navn):
    postnummer = df.iloc[6, 1:].dropna().values
    mengder = df.iloc[8, 1:].dropna().values
    kommentar = f"{objekt_navn}: Applag"
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
    return pd.DataFrame(data)

def process_sfi_sheet(df, objekt_navn):
    postnummer = df.iloc[5, 1:].values
    mengder = df.iloc[14, 1:].values
    kommentar = f"{objekt_navn}: Lengdeprofil"
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
    return pd.DataFrame(data)

def process_xfi_sheet(df, objekt_navn):
    postnummer = df.iloc[8:, 0].dropna().values
    mengder = df.iloc[8:, 10].dropna().values
    kommentar = f"{objekt_navn}: XFI"
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
    return pd.DataFrame(data)

def process_efi_sheet(df, objekt_navn):
    postnummer = df.iloc[8:, 0].dropna().values
    mengder = df.iloc[8:, 10].dropna().values
    kommentar = f"{objekt_navn}: EFI"
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
    return pd.DataFrame(data)

st.title("Excel-filbehandling med dokumentasjonskobling")
excel_file = st.file_uploader("Last opp en Excel-fil", type=["xlsx", "xls", "xlsm"])

if excel_file:
    xl = pd.ExcelFile(excel_file)
    for sheet_name in xl.sheet_names:
        objekt_navn = sheet_name.split('.')[0]
        if sheet_name.endswith('.aly'):
            df = xl.parse(sheet_name)
            processed_df = process_aly_sheet(df, objekt_navn)
        elif sheet_name.endswith('.sfi'):
            df = xl.parse(sheet_name)
            processed_df = process_sfi_sheet(df, objekt_navn)
        elif sheet_name.endswith('.xfi'):
            df = xl.parse(sheet_name)
            processed_df = process_xfi_sheet(df, objekt_navn)
        elif sheet_name.endswith('.efi'):
            df = xl.parse(sheet_name)
            processed_df = process_efi_sheet(df, objekt_navn)
        else:
            continue

        st.write(f"Behandlet data fra arkfane: {sheet_name}")
        st.dataframe(processed_df)

        # Lagre resultatet i en buffer
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            processed_df.to_excel(writer, index=False, sheet_name=sheet_name)
        buffer.seek(0)

        # Provide the file for download
        st.download_button(
            label="Download processed file",
            data=buffer,
            file_name=f"processed_{sheet_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
