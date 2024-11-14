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

if excel_file:
    xl = pd.ExcelFile(excel_file)
    
    # Process each sheet individually
    for sheet_name in xl.sheet_names:
        objekt_navn = sheet_name.split('.')[0]
        df = xl.parse(sheet_name)
        
        # Only prompt for file path if sheet name ends with specified extensions
        if sheet_name.endswith(('.aly', '.sfi', '.xfi', '.efi')):
            st.write(f"Velg dokumentasjonsfil for ark: {sheet_name}")
            link_file = st.file_uploader(f"Last opp lenkefil for {sheet_name}", key=sheet_name)

            if link_file is not None:
                # Use the uploaded file's path as a hyperlink in the Excel table
                link = link_file.name

                # Process based on sheet type
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

                # Display the processed data
                st.write(f"Behandlet data fra arkfane: {sheet_name}")
                st.dataframe(processed_df)

                # Prepare a downloadable file for each sheet
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    processed_df.to_excel(writer, index=False, sheet_name=sheet_name)
                    workbook = writer.book
                    worksheet = writer.sheets[sheet_name]
                    
                    # Format as an Excel table
                    max_row, max_col = processed_df.shape
                    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': [{'header': col} for col in processed_df.columns]})
                
                buffer.seek(0)

                # Download button for each sheet
                st.download_button(
                    label=f"Last ned behandlet fil for {sheet_name}",
                    data=buffer,
                    file_name=f"behandlet_{sheet_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
