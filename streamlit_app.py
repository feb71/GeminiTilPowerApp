import pandas as pd
import streamlit as st
import io

def process_aly_sheet(df, objekt_navn):
    # Henter postnummer fra rad 7, kolonne B og utover
    postnummer = df.iloc[5, 1:].dropna().values  # Rad 7 tilsvarer indeks 6
    # Henter mengder fra rad 9, kolonne B og utover
    mengder = df.iloc[7, 1:].dropna().values  # Rad 9 tilsvarer indeks 8
    # Oppretter en kommentar basert på objektets navn
    kommentar = [f"{objekt_navn}: Applag"] * len(postnummer)
    # Kombinerer dataene i en DataFrame
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
    return pd.DataFrame(data)

def process_sfi_cross_section(df, objekt_navn):
    # Behandlingslogikk for tverrprofil
    postnummer = df.iloc[5, 1:].dropna().values
    mengder = df.iloc[14, 1:].dropna().values
    profiler = df.iloc[16:, 0].dropna().values
    første_profil = profiler[0] if len(profiler) > 0 else None
    siste_profil = profiler[-1] if len(profiler) > 0 else None
    kommentar = [f"{objekt_navn}: Fra profil {første_profil} til profil {siste_profil}"] * len(postnummer) if første_profil and siste_profil else [f"{objekt_navn}: Tverrprofil"] * len(postnummer)
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
    return pd.DataFrame(data)

def process_sfi_longitudinal(df, objekt_navn):
    # Behandlingslogikk for lengdeprofil
    postnummer = df.iloc[5, 1:].dropna().values  # Postnummer i rad 7, fra kolonne B
    mengder = df.iloc[14, 1:].dropna().values  # Mengde i rad 16, fra kolonne B
    kommentar = [f"{objekt_navn}: Lengdeprofil"] * len(postnummer)  # Enklere kommentar for lengdeprofil
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
    return pd.DataFrame(data)

def determine_sfi_type(df):
    # Sjekker innholdet i rad 18, kolonne A
    if df.shape[0] > 16 and df.shape[1] > 0:
        cell_value = str(df.iloc[16, 0]).strip().lower()
        if cell_value == 'l':
            return "longitudinal"
    
    # Sjekker enheter i rad 12, fra kolonne B og utover
    if df.shape[0] > 10 and df.shape[1] > 1:
        units = df.iloc[10, 1:].dropna().astype(str).str.strip().str.lower()
        if units.isin(['m²', 'm³', 'm2', 'm3']).any():
            return "cross_section"
        elif units.isin(['m']).all():
            return "longitudinal"
    
    # Standard til tverrprofil hvis ingen kriterier er oppfylt
    return "cross_section"

def process_xfi_sheet(df, objekt_navn):
    # Tilpasset behandling for .xfi-arkfaner
    postnummer = df.iloc[7:, 0].dropna().values  # Postnummer fra rad 9, kolonne A
    mengder = df.iloc[7:, 10].dropna().values  # Mengder fra rad 9, kolonne K
    kommentar = [f"{objekt_navn}: XFI"] * len(postnummer)
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
    return pd.DataFrame(data)

def process_efi_sheet(df, objekt_navn):
    # Tilpasset behandling for .efi-arkfaner
    postnummer = df.iloc[7:, 0].dropna().values  # Postnummer fra rad 9, kolonne A
    mengder = df.iloc[7:, 10].dropna().values  # Mengder fra rad 9, kolonne K
    kommentar = [f"{objekt_navn}: EFI"] * len(postnummer)
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
    return pd.DataFrame(data)

# Streamlit-oppsett
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
