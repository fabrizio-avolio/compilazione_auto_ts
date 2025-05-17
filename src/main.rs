use internal_resources::{calendario, file_handling,utility_worksheet};
use rust_xlsxwriter::{Workbook, XlsxError};
mod internal_resources;
fn main() -> Result<(), XlsxError> {
    let calendario = calendario::Calendario::new();

    //------Creazione variabili------//
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let nome_file = file_handling::get_file_path(calendario.anno, calendario.mese);
    //------Utility per celle custom ------//
    utility_worksheet::custom_height_and_width_for_cells(worksheet);

    //------Stringhe Statiche ------//
    utility_worksheet::build_static_strings_in_excel(worksheet);

    //------Stringhe dinamiche Nominativo e Cliente ------//
    utility_worksheet::write_nominativo_and_cliente(worksheet);

    // -----Costruzione mese e anno ------//
    utility_worksheet::build_month_and_year(worksheet, &calendario);

    //-----Costruzione Dinamica del mese----//
    utility_worksheet::costruisci_gg_nome_gg(worksheet, &calendario);

    //----- Costruzione Della prima intestazione------//
    utility_worksheet::build_top_lines(worksheet);

    //----- Save e finestra di dialogo------//
    file_handling::safe_save_workbook(&mut workbook, &nome_file);
    file_handling::file_creato();

    Ok(())
}
