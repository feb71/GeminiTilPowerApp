import pandas as pd
import streamlit as st
import io
from urllib.parse import quote

def process_aly_sheet(df, objekt_navn):
    try:
        postnummer = df.iloc[5, 1:].dropna().values
        mengder = df.iloc[7, 1:].dropna().values
        kommentar = [f"{objekt_navn}: Applag"] * len(postnummer)
        data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
        return pd.DataFrame(data)
    except IndexError:
        st.error(f"Feil ved behandling av {objekt_navn}: Sjekk at arket har riktig format.")
        return pd.DataFrame()

def process_sfi_cross_section(df, objekt_navn):
    try:
        postnummer = df.iloc[5, 1:].dropna().values
        mengder = df.iloc[14, 1:].dropna().values
        profiler = df.iloc[16:, 0].dropna().values
        første_profil = profiler[0] if len(profiler) > 0 else "0.000"
        siste_profil = profiler[-1] if len(profiler) > 0 else "0.000"
        kommentar = [f"{objekt_navn}: Fra profil {første_profil} til profil {siste_profil}"] * len(postnummer)
        data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
        return pd.DataFrame(data)
    except IndexError:
        st.error(f"Feil ved behandling av {objekt_navn}: Sjekk at arket har riktig format.")
        return pd.DataFrame()

def process_sfi_longitudinal(df, objekt_navn):
    try:
        postnummer = df.iloc[5, 1:].dropna().values
        mengder = df.iloc[14, 1:].dropna().values
        kommentar = [f"{objekt_navn}: Lengdeprofil"] * len(postnummer)
        data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
        return pd.DataFrame(data)
    except IndexError:
        st.error(f"Feil ved behandling av {objekt_navn}: Sjekk at arket har riktig format.")
        return pd.DataFrame()

def determine_sfi_type(df):
    if df.shape[0] > 16 and df.shape[1] > 0:
        cell_value = str(df.iloc[16, 0]).strip().lower()
        if cell_value == 'l':
            return "longitudinal"
    if df.shape[0] > 10 and df.shape[1] > 1:
        units = df.iloc[10, 1:].dropna().astype(str).str.strip().str.lower()
        if units.isin(['m²', 'm³', 'm2', 'm3']).any():
            return "cross_section"
        elif units.isin(['m']).all():
            return "longitudinal"
    return "cross_section"

def process_xfi_sheet(df, objekt_navn):
    try:
        postnummer = df.iloc[7:, 0].dropna().values
        mengder = df.iloc[7:, 10].dropna().values
        kommentar = [f"{objekt_navn}: XFI"] * len(postnummer)
        data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
        return pd.DataFrame(data)
    except IndexError:
        st.error(f"Feil ved behandling av {objekt_navn}: Sjekk at arket har riktig format.")
        return pd.DataFrame()

def process_efi_sheet(df, objekt_navn):
    try:
        postnummer = df.iloc[7:, 0].dropna().values
        mengder = df.iloc[7:, 10].dropna().values
        kommentar = [f"{objekt_navn}: EFI"] * len(postnummer)
        data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
        return pd.DataFrame(data)
    except IndexError:
        st.error(f"Feil ved behandling av {objekt_navn}: Sjekk at arket har riktig format.")
        return pd.DataFrame()

def create_hyperlink(file_name, sharepoint_url):
    encoded_file_name = quote(file_name)
    return f"{sharepoint_url}/{encoded_file_name}"

# Streamlit-oppsett
st.title("Excel-filbehandling med dokumentasjonskobling")
excel_file = st.file_uploader("Last opp en Excel-fil", type=["xlsx", "xls", "xlsm"])

if excel_file:
    xl = pd.ExcelFile(excel_file)
    sharepoint_url = st.text_input("Skriv inn SharePoint-bibliotekets URL:")

    for sheet_name in xl.sheet_names:
        objekt_navn = sheet_name.split('.')[0]
        if sheet_name.endswith('.aly'):
            df = xl.parse(sheet_name)
            processed_df = process_aly_sheet(df, objekt_navn)
        elif sheet_name.endswith('.sfi'):
            df = xl.parse(sheet_name)
            sfi_type = determine_sfi_type(df)
            if sfi_type == "cross_section":
                processed_df = process_sfi_cross_section(df, objekt_navn)
            else:
                processed_df = process_sfi_longitudinal(df, objekt_navn)
        elif sheet_name.endswith('.xfi'):
            df = xl.parse(sheet_name)
            processed_df = process_xfi_sheet(df, objekt_navn)
        elif sheet_name.endswith('.efi'):
            df = xl.parse(sheet_name)
            processed_df = process_efi_sheet(df, objekt_navn)
        else:
            continue  # Hopp over arkfaner som ikke matcher

        if not processed_df.empty:
            if sharepoint_url:
                processed_df['Dokumentasjon'] = processed_df['Postnummer'].apply(
                    lambda x: create_hyperlink(f"{objekt_navn}_{x}.pdf", sharepoint_url)
                )
            st.write(f"Behandlet data fra arkfane: {sheet_name}")
            st.dataframe(processed_df)

            # Lagre resultatet i en buffer
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                processed_df.to_excel(writer, index=False, sheet_name=sheet_name)
            buffer.seek(0)

            # Tilby nedlasting av filen
            st.download_button(
                label=f"Last ned behandlet fil for {sheet_name}",
                data=buffer,
                file_name=f"behandlet_{sheet_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
