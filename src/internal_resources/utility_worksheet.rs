
use std::str::FromStr;

use chrono::{NaiveTime, Timelike};
use rust_xlsxwriter::{Format, Worksheet, XlsxBorder,XlsxAlign};

use crate::internal_resources::formats::times_new_yellow;

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


    
    for i in 0..15 {
    match i {
        2 => safe_writing_strings_in_cells(worksheet, 7, i, "         Orario Entrata-Uscita", &formats::top_bordered_medium_bold()),
        14 => safe_writing_blank_in_cells(worksheet, 7, i, &rust_xlsxwriter::Format::new().set_border_top(XlsxBorder::Medium)),
        _ => safe_writing_blank_in_cells(worksheet, 7, i, &formats::top_bordered_medium_bold()),
    }
}





    for i in 0..14 {
    safe_writing_blank_in_cells(worksheet,9 , i, &Format::new().set_border_bottom(XlsxBorder::Medium));
    }

    let intestazioni: [&str; 13] = ["Mattina","","Pomeriggio","","Ore", "Straor.", "Ferie","Perm","Fest.","Recupero","Malat","Sede","N O T E"];
    let mut i :u16= 2;
    for intestazione in intestazioni{
        match i {
    14 => {
        safe_writing_strings_in_cells(worksheet,8,i,intestazione,&formats::centered_text_bold());
    }
    _ => {
        safe_writing_strings_in_cells(worksheet, 8, i, intestazione, &formats::borded_bold_text(10));
    }
}
    i+=1;
    }

    safe_writing_strings_in_cells(worksheet,8,6,"Ore", &rust_xlsxwriter::Format::new()
    .set_bold()
    .set_font_name("Times New Roman")
    .set_font_size(10)
    .set_border_top(XlsxBorder::Thin)
    .set_border_bottom(XlsxBorder::Thin));

    safe_writing_strings_in_cells(worksheet,8,7,"Straor.", &rust_xlsxwriter::Format::new()
    .set_bold()
    .set_font_name("Times New Roman")
    .set_font_size(10)
    .set_border_top(XlsxBorder::Thin)
    .set_border_bottom(XlsxBorder::Thin));

    
    
}
// ---------------------------------------------------------------------------------//
pub fn colored_yellow_if_checked(worksheet: &mut Worksheet, row: u32, col: u16, string: &str) {
    match string {
        "sabato" | "domenica" => {
            safe_writing_strings_in_cells(worksheet,row, col, string, &formats::times_new_yellow());
            for k in 2..15{
                safe_writing_blank_in_cells(worksheet, row, k, &formats::times_new_yellow());
                if k==14{
                    safe_writing_blank_in_cells(worksheet, row, k, &formats::times_new_yellow().set_border_right(XlsxBorder::Medium))
                }
            }
        }
        _ => {
            safe_writing_strings_in_cells(worksheet,row, col, string,&rust_xlsxwriter::Format::new().set_border_bottom(XlsxBorder::Thin).set_border_top(XlsxBorder::Thin).set_border_right(XlsxBorder::Thin));
        }
    }
// ---------------------------------------------------------------------------------//
}
// ---------------------------------------------------------------------------------//
pub fn build_top_lines (worksheet:&mut Worksheet) {
    let top_intestazioni: [&str; 15] = ["GG", "", "Ent.","Usc.","Ent.","Usc.","Tot. gg","","","","","","","",""];
    let mut l = 0;
    for top_intestazione in top_intestazioni {
        match l {
           0 => {
            print!("{}", top_intestazione);
            safe_writing_strings_in_cells(worksheet, 9, l, top_intestazione, &formats::times_new_yellow_centered_bold(8));
           }
           1=> {
            safe_writing_strings_in_cells(worksheet, 9, l, top_intestazione,&formats::borded_bold_text_bottom_medium (10));
           }
           14=>{
            let format = rust_xlsxwriter::Format::new().set_font_name("Times New Roman").set_bold().set_font_size(10).set_border_bottom(XlsxBorder::Medium).set_border_left(XlsxBorder::Thin);
            safe_writing_strings_in_cells(worksheet, 9 , l, top_intestazione, &format);

           }
           _=> {
            safe_writing_strings_in_cells(worksheet, 9 , l, top_intestazione, &formats::borded_bold_text_bottom_medium (10));
           }
        }
        l+=1
    }
}
// ---------------------------------------------------------------------------------//
fn time_to_excel(time_str: &str) -> Result<f64, chrono::ParseError> {
    let time = NaiveTime::from_str(time_str)?;
    let hours = time.hour() as f64;
    let minutes = time.minute() as f64;
    Ok((hours + minutes / 60.0) / 24.0)
}

fn calculate_time_difference(first_cell: f64, second_cell: f64, third_cell: f64, fourth_cell: f64) -> String {
    let result = if first_cell == 0.0 {
        if second_cell == 0.0 {
            0.0
        } else {
            second_cell - third_cell
        }
    } else {
        if second_cell == 0.0 {
            first_cell - fourth_cell
        } else {
            (first_cell - fourth_cell) + (second_cell - third_cell)
        }
    };

    let result_in_hours = result * 24.0;

    if result_in_hours == 0.0 {
        " ".to_string()
    } else {
        result_in_hours.to_string()
    }
}