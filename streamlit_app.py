import pandas as pd
import streamlit as st
import io

# Define processing functions with hyperlinks
def process_aly_sheet(df, objekt_navn, link):
    postnummer = df.iloc[5, 1:].dropna().values
    mengder = df.iloc[7, 1:].dropna().values
    kommentar = [f"{objekt_navn}: Applag"] * len(postnummer)
    hyperlinks = [f'=HYPERLINK("{link}", "Dokumentasjon")'] * len(postnummer)
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar, "Link": hyperlinks}
    return pd.DataFrame(data)

def process_sfi_cross_section(df, objekt_navn, link):
    postnummer = df.iloc[5, 1:].dropna().values
    mengder = df.iloc[14, 1:].dropna().values
    profiler = df.iloc[16:, 0].dropna().values
    første_profil = profiler[0] if len(profiler) > 0 else None
    siste_profil = profiler[-1] if len(profiler) > 0 else None
    kommentar = [f"{objekt_navn}: Fra profil {første_profil} til profil {siste_profil}"] * len(postnummer) if første_profil is not None and siste_profil is not None else [f"{objekt_navn}: Tverrprofil"] * len(postnummer)
    hyperlinks = [f'=HYPERLINK("{link}", "Dokumentasjon")'] * len(postnummer)
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar, "Link": hyperlinks}
    return pd.DataFrame(data)

def process_sfi_longitudinal(df, objekt_navn, link):
    postnummer = df.iloc[5, 1:].dropna().values
    mengder = df.iloc[14, 1:].dropna().values
    kommentar = [f"{objekt_navn}: Lengdeprofil"] * len(postnummer)
    hyperlinks = [f'=HYPERLINK("{link}", "Dokumentasjon")'] * len(postnummer)
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar, "Link": hyperlinks}
    return pd.DataFrame(data)

def process_xfi_sheet(df, objekt_navn, link):
    postnummer = df.iloc[7:, 0].dropna().values
    mengder = df.iloc[7:, 10].dropna().values
    kommentar = [f"{objekt_navn}: XFI"] * len(postnummer)
    hyperlinks = [f'=HYPERLINK("{link}", "Dokumentasjon")'] * len(postnummer)
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar, "Link": hyperlinks}
    return pd.DataFrame(data)

def process_efi_sheet(df, objekt_navn, link):
    postnummer = df.iloc[7:, 0].dropna().values
    mengder = df.iloc[7:, 10].dropna().values
    kommentar = [f"{objekt_navn}: EFI"] * len(postnummer)
    hyperlinks = [f'=HYPERLINK("{link}", "Dokumentasjon")'] * len(postnummer)
    data = {"Postnummer": postnummer, "Mengde": mengder, "Kommentar": kommentar, "Link": hyperlinks}
    return pd.DataFrame(data)

# Streamlit setup
st.title("Excel-filbehandling med dokumentasjonskobling")
excel_file = st.file_uploader("Last opp en Excel-fil", type=["xlsx", "xls", "xlsm"])

# Input for document link
link = st.text_input("Skriv inn lenken til dokumentasjonen som gjelder for hvert objekt:")

if excel_file and link:
    xl = pd.ExcelFile(excel_file)
    processed_sheets = {}

    for sheet_name in xl.sheet_names:
        objekt_navn = sheet_name.split('.')[0]
        df = xl.parse(sheet_name)
        
        # Process each sheet with hyperlink
        if sheet_name.endswith('.aly'):
            processed_df = process_aly_sheet(df, objekt_navn, link)
        elif sheet_name.endswith('.sfi'):
            sfi_type = determine_sfi_type(df)
            if sfi_type == "cross_section":
                processed_df = process_sfi_cross_section(df, objekt_navn, link)
            else:
                processed_df = process_sfi_longitudinal(df, objekt_navn, link)
        elif sheet_name.endswith('.xfi'):
            processed_df = process_xfi_sheet(df, objekt_navn, link)
        elif sheet_name.endswith('.efi'):
            processed_df = process_efi_sheet(df, objekt_navn, link)
        else:
            continue

        st.write(f"Behandlet data fra arkfane: {sheet_name}")
        st.dataframe(processed_df)
        
        # Store processed DataFrame with hyperlink column
        processed_sheets[sheet_name] = processed_df

    # Save all processed DataFrames to a single Excel file with table formatting
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        for sheet, data in processed_sheets.items():
            data.to_excel(writer, index=False, sheet_name=sheet)
            workbook = writer.book
            worksheet = writer.sheets[sheet]

            # Get the dimensions of the DataFrame
            max_row, max_col = data.shape

            # Define the table range and add it to the sheet as a table
            worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': [{'header': col} for col in data.columns]})
    buffer.seek(0)

    # Download button for combined file with tables
    st.download_button(
        label="Last ned samlet behandlet fil med hyperkoblinger og tabellformat",
        data=buffer,
        file_name="behandlet_filer_med_tabel.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
