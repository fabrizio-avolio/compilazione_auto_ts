
use rust_xlsxwriter::{Format, Worksheet, XlsxBorder,XlsxAlign};

use super::calendario::Calendario;
use super::formats;

//-------CONST-------//
const SIZE_ERROR: &str = "Errore impostando la larghezza: ";
const WRITING_ERROR: &str = "Errore nella scrittura su cella: ";

// ---------------------------------------------------------------------------------//
pub(crate) fn costruisci_gg_nome_gg(worksheet: &mut Worksheet, calendario: &Calendario) {
    let mut i:u32 = 10;
    for giorno in &calendario.giorni {
        if i == 0 {
            safe_writing_strings_in_cells (worksheet,10,0,&giorno.giorno.to_string(),&formats::times_new_centered());
            colored_yellow_if_checked(worksheet,10,1,&giorno.giorno_settimana.to_string());
            
        } else {
            safe_writing_strings_in_cells (worksheet,i,0,&giorno.giorno.to_string(),&formats::times_new_centered());
            colored_yellow_if_checked(worksheet,i,1,&giorno.giorno_settimana.to_string());
        }
        i += 1;
    }
}
// ---------------------------------------------------------------------------------//
pub(crate) fn custom_height_and_width_for_cells(worksheet: &mut Worksheet) {
    safe_method_to_modify_cell_width(worksheet, 0, 3.67);
    safe_method_to_modify_cell_height(worksheet, 2, 8.0);
    safe_method_to_modify_cell_height(worksheet, 3, 0.0);
    safe_method_to_modify_cell_height(worksheet, 6, 8.0);
    safe_method_to_modify_cell_height(worksheet, 5, 20.75);
    safe_method_to_modify_cell_width(worksheet, 14, 70.86);
    safe_method_to_modify_cell_width(worksheet, 1, 9.15);
}
// ---------------------------------------------------------------------------------//
fn safe_method_to_modify_cell_width(worksheet: &mut Worksheet, col: u16, w: f64) {
    match worksheet.set_column_width(col, w) {
        Ok(_) => (),
        Err(e) => eprintln!("{} {}", SIZE_ERROR, e),
    }
}
// ---------------------------------------------------------------------------------//
fn safe_method_to_modify_cell_height(worksheet: &mut Worksheet, row: u32, h: f64) {
    match worksheet.set_row_height(row, h) {
        Ok(_) => (),
        Err(e) => eprintln!("{} {}", SIZE_ERROR, e),
    }
}
// ---------------------------------------------------------------------------------//
fn safe_writing_strings_in_cells (worksheet: &mut Worksheet,row: u32,col: u16,string:&str,format:&Format) {
    match worksheet.write_string(row, col, string, format){
        Ok(_) => (),
        Err(e) => eprintln!("{} {}", WRITING_ERROR, e),
    }
}
// ---------------------------------------------------------------------------------//
fn safe_writing_blank_in_cells (worksheet: &mut Worksheet,row: u32,col:u16,format:&Format) {
    match worksheet.write_blank(row, col, format){
        Ok(_) => (),
        Err(e) => eprintln!("{} {}", WRITING_ERROR, e),
    }
}
// ---------------------------------------------------------------------------------//



pub fn build_static_strings_in_excel (worksheet: &mut Worksheet) {

    safe_writing_strings_in_cells(worksheet, 1, 3, "Modulo rilevazione presenza del personale", &formats::times_new_roman_bold_italic());

    safe_writing_strings_in_cells(worksheet, 4, 3, "Cliente:", &formats::times_new_roman_italic());

    safe_writing_strings_in_cells(worksheet,5,3,"Nominativo",&formats::times_new_roman_italic());

    safe_writing_strings_in_cells(worksheet,1,12,"Mese:",&formats::times_new_roman_italic());

    safe_writing_strings_in_cells(worksheet,7,2,"         Orario Entrata-Uscita", &rust_xlsxwriter::Format::new()
    .set_bold()
    .set_font_name("Times New Roman")
    .set_font_size(10)
    .set_border_top(XlsxBorder::Medium));

    safe_writing_strings_in_cells(worksheet,8,2,"      Mattina", &rust_xlsxwriter::Format::new()
    .set_bold()
    .set_font_name("Times New Roman")
    .set_font_size(10)
    .set_border_top(XlsxBorder::Medium));

    safe_writing_strings_in_cells(worksheet,8,4,"    Pomeriggio", &rust_xlsxwriter::Format::new()
    .set_bold()
    .set_font_name("Times New Roman")
    .set_font_size(10)
    .set_border_top(XlsxBorder::Medium));

    //safe_writing_strings_in_cells(worksheet,9,0,"GG",)

    for i in 0..14 {
    safe_writing_blank_in_cells(worksheet,9 , i, &Format::new().set_border_bottom(XlsxBorder::Medium));
    }

    let intestazioni: [&str; 9] = ["Ore", "Straor.", "Ferie","Perm","Fest.","Recupero","Malat","Sede","N O T E"];
    let mut i :u16= 6;
    for intestazione in intestazioni{
        if i==14{
            safe_writing_strings_in_cells(worksheet,8,i,intestazione, &rust_xlsxwriter::Format::new()
            .set_align(XlsxAlign::Center)
            .set_bold()
            .set_font_name("Times New Roman")
            .set_font_size(10)
            .set_border_top(XlsxBorder::Thin)
            .set_border_left(XlsxBorder::Thin)
            .set_border_bottom(XlsxBorder::Thin)
            .set_border_right(XlsxBorder::Thin));
        }
        else{
            safe_writing_strings_in_cells(worksheet,8,i,intestazione, &rust_xlsxwriter::Format::new()
                .set_bold()
                .set_font_name("Times New Roman")
                .set_font_size(10)
                .set_border_top(XlsxBorder::Thin)
                .set_border_left(XlsxBorder::Thin)
                .set_border_bottom(XlsxBorder::Thin)
                .set_border_right(XlsxBorder::Thin));
        }
    i+=1;
    }

    safe_writing_strings_in_cells(worksheet,8,6,"Ore", &rust_xlsxwriter::Format::new()
    .set_bold()
    .set_font_name("Times New Roman")
    .set_font_size(10)
    .set_border_top(XlsxBorder::Medium));

    safe_writing_strings_in_cells(worksheet,8,7,"Straor.", &rust_xlsxwriter::Format::new()
    .set_bold()
    .set_font_name("Times New Roman")
    .set_font_size(10)
    .set_border_top(XlsxBorder::Medium));

    
    
}
// ---------------------------------------------------------------------------------//
pub fn colored_yellow_if_checked(worksheet: &mut Worksheet, row: u32, col: u16, string: &str) {
    match string {
        "sabato" | "domenica" => {
            safe_writing_strings_in_cells(worksheet,row, col, string, &formats::times_new_yellow());
            for k in 2..14{
                safe_writing_blank_in_cells(worksheet, row, k, &formats::times_new_yellow());
            }
        }
        _ => {
            safe_writing_strings_in_cells(worksheet,row, col, string,&rust_xlsxwriter::Format::new().set_border_bottom(XlsxBorder::Thin).set_border_top(XlsxBorder::Thin).set_border_right(XlsxBorder::Thin));
        }
    }
// ---------------------------------------------------------------------------------//

}