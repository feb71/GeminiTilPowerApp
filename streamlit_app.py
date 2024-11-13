import pandas as pd
import streamlit as st
import os

# Synkronisert mappe for lagring
output_folder = "path/to/your/synced/folder"

def process_aly_sheet(df, objekt_navn):
    # Henter postnummer fra rad 7, kolonne B og utover
    postnummer = df.iloc[6, 1:].dropna().values  # Rad 7 tilsvarer indeks 6

    # Henter mengder fra rad 9, kolonne B og utover
    mengder = df.iloc[8, 1:].dropna().values  # Rad 9 tilsvarer indeks 8

    # Oppretter en kommentar basert på objektets navn
    kommentar = f"{objekt_navn}: Applag"

    # Kombinerer dataene i en DataFrame
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
    return pd.DataFrame(data)

def process_sfi_sheet(df, objekt_navn):
    # Tilpasset behandling for .sfi-arkfaner
    # Implementer logikk basert på .sfi-strukturen
    # Eksempel:
    postnummer = df.iloc[6, 1:].values  # Postnummer i rad 7, fra kolonne B
    mengder = df.iloc[15, 1:].values  # Mengder i rad 16, fra kolonne B
    kommentar = f"{objekt_navn}: Lengdeprofil"
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
    return pd.DataFrame(data)

def process_xfi_sheet(df, objekt_navn):
    # Tilpasset behandling for .xfi-arkfaner
    # Implementer logikk basert på .xfi-strukturen
    # Eksempel:
    postnummer = df.iloc[8:, 0].dropna().values  # Postnummer fra rad 9, kolonne A
    mengder = df.iloc[8:, 10].dropna().values  # Mengder fra rad 9, kolonne K
    kommentar = f"{objekt_navn}: XFI"
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar}
    return pd.DataFrame(data)

def process_efi_sheet(df, objekt_navn):
    # Tilpasset behandling for .efi-arkfaner
    postnummer = df.iloc[8:, 0].dropna().values  # Postnummer fra rad 9, kolonne A
    mengder = df.iloc[8:, 10].dropna().values  # Mengder fra rad 9, kolonne K
    kommentar = f"{objekt_navn}: EFI"
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
            processed_df = process_sfi_sheet(df, objekt_navn)
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
        
        # Lagre resultatet
        save_path = os.path.join(output_folder, f"behandlet_{sheet_name}.xlsx")
        processed_df.to_excel(save_path, index=False)
        st.success(f"Resultatet er lagret i {save_path}")
