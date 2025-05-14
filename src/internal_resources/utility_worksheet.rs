use rust_xlsxwriter::{Format, FormatBorder, Worksheet};

use super::calendario::Calendario;
use super::formats;

//-------CONST-------//
const SIZE_ERROR: &str = "Errore impostando la larghezza: ";
const WRITING_ERROR: &str = "Errore nella scrittura su cella: ";

// ---------------------------------------------------------------------------------//
pub(crate) fn costruisci_gg_nome_gg(worksheet: &mut Worksheet, calendario: &Calendario) {
    let mut i:u32= 10;
    for giorno in &calendario.giorni {
            safe_writing_strings_in_cells (worksheet,i,0,&giorno.giorno.to_string(),&formats::times_new_centered());
            colored_yellow_if_checked(worksheet,i,1,&giorno.giorno_settimana.to_string());
        i += 1;
    }
    let formule: [&str; 8] =["=SUM(G11:G{})","=SUM(H11:H{})","=SUM(I11:I{})","=SUM(J11:J{})","=SUM(K11:K{})","=SUM(L11:L{})","=SUM(M11:M{})","=SUM(N11:N{})"]; 
    let mut k = 0;
    for formula in formule{
        let new_formula = formula.replace("{}", &i.to_string());
        print!("new_formula{} ", new_formula);
        match k {
            0 => {
                let _ = worksheet.write_formula_with_format(i as u32, (6+k).try_into().unwrap(), &rust_xlsxwriter::Formula::new(&new_formula), &formats::times_new_green_borded().set_bold());
            },
            _ => {
                let _ = worksheet.write_formula_with_format(i as u32, (6+k).try_into().unwrap(), &rust_xlsxwriter::Formula::new(&new_formula), &formats::borded_bold_text(10));
            }
        }
        k+=1;
    }
   let somme: [&str; 8] = ["G{}","H{}","I{}","J{}","K{}","L{}","M{}","N{}"];
   let mut somma_tot:String = String::new(); 
    for cella in somme {
        somma_tot.push_str("+");
        let updated_cella = cella.replace("{}", &(i+1).to_string());
        somma_tot.push_str(&updated_cella);
    }
    somma_tot.insert(0, '=');
    print!("somma_tot {}", somma_tot);
   safe_method_to_modify_cell_height(worksheet, i+1, 9.0);
   safe_writing_strings_in_cells(worksheet, (i+2).try_into().unwrap(), 0, "Totale ore di presenza nel mese :", &formats::times_new_roman_italic_with_font(13));
   let _ = worksheet.write_formula_with_format((i+2).try_into().unwrap(), 5, &rust_xlsxwriter::Formula::new(&somma_tot), &formats::times_new_centered_bolded());
   
   
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
    match worksheet.write_string_with_format(row, col, string, format){
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
        14 => safe_writing_blank_in_cells(worksheet, 7, i, &rust_xlsxwriter::Format::new().set_border_top(FormatBorder::Medium).set_border_right(FormatBorder::Medium)),
        _ => safe_writing_blank_in_cells(worksheet, 7, i, &formats::top_bordered_medium_bold()),
    }
}





    for i in 0..14 {
    safe_writing_blank_in_cells(worksheet,9 , i, &Format::new().set_border_bottom(FormatBorder::Medium));
    }

    let intestazioni: [&str; 13] = ["Mattina","","Pomeriggio","","Ore", "Straor.", "Ferie","Perm","Fest.","Recupero","Malat","Sede","N O T E"];
    let mut i :u16= 2;
    for intestazione in intestazioni{
        match i {
    14 => {
        safe_writing_strings_in_cells(worksheet,8,i,intestazione,&formats::centered_text_bold().set_border_right(FormatBorder::Medium));
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
    .set_border_top(FormatBorder::Thin)
    .set_border_bottom(FormatBorder::Thin));

    safe_writing_strings_in_cells(worksheet,8,7,"Straor.", &rust_xlsxwriter::Format::new()
    .set_bold()
    .set_font_name("Times New Roman")
    .set_font_size(10)
    .set_border_top(FormatBorder::Thin)
    .set_border_bottom(FormatBorder::Thin));

    
    
}
// ---------------------------------------------------------------------------------//
pub fn colored_yellow_if_checked(worksheet: &mut Worksheet, row: u32, col: u16, string: &str) {
    match string {
        "sabato" | "domenica" => {
            safe_writing_strings_in_cells(worksheet,row, col, string, &formats::times_new_yellow());
            for k in 2..15{
                safe_writing_blank_in_cells(worksheet, row, k, &formats::times_new_yellow());
                if k==14{
                    safe_writing_blank_in_cells(worksheet, row, k, &formats::times_new_yellow().set_border_right(FormatBorder::Medium))
                }
            }
        }
        _ => {
            safe_writing_strings_in_cells(worksheet,row, col, string,&rust_xlsxwriter::Format::new().set_border_bottom(FormatBorder::Thin).set_border_top(FormatBorder::Thin).set_border_right(FormatBorder::Thin));
            safe_writing_strings_in_cells(worksheet,row, col+1, "08:30:00", &formats::borded_normal_text(10));
            safe_writing_strings_in_cells(worksheet,row, col+2, "13:00:00", &formats::borded_normal_text(10));
            safe_writing_strings_in_cells(worksheet,row, col+3, "14:00:00", &formats::borded_normal_text(10));
            safe_writing_strings_in_cells(worksheet,row, col+4, "17:30:00", &formats::borded_normal_text(10));
            let formula = generate_formula(row, col+5);
            let formula_converted = rust_xlsxwriter::Formula::new(formula);
            let _ = worksheet.write_formula_with_format(row, col+5 , formula_converted,&formats::borded_bold_text(10));
            for k in 7..15{
                safe_writing_blank_in_cells(worksheet, row, k, &formats::borded_normal_text(10));
                if k==14{
                    safe_writing_blank_in_cells(worksheet, row, k, &formats::borded_no_right().set_border_right(FormatBorder::Medium))
                }
            }
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
            let format = rust_xlsxwriter::Format::new().set_font_name("Times New Roman").set_bold().set_font_size(10).set_border_bottom(FormatBorder::Medium).set_border_left(FormatBorder::Thin).set_border_right(FormatBorder::Medium);
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
fn translate_cell(row: u32, col: u16) -> String {
    let mut col_str = String::new();
    let mut col_num = col;

    while col_num < 26 {
        col_str.insert(0, (b'A' + (col_num % 26) as u8) as char);
        if col_num >= 26 {
            col_num = col_num / 26 - 1;
        } else {
            break;
        }
    }

    format!("{}{}", col_str, row + 1)
}
// ---------------------------------------------------------------------------------//
fn generate_formula(row: u32, col: u16) -> String {
    let _cell_ref = translate_cell(row, col);
    let cell_ref_1 = translate_cell(row, col - 1);
    let cell_ref_2 = translate_cell(row, col - 2);
    let cell_ref_3 = translate_cell(row, col - 3);
    let cell_ref_4 = translate_cell(row, col - 4);

    format!(
        "=IF(IF({}=0,IF({}=0,\"0\",{}-{}),IF({}=0,{}-{},({}-{})+({}-{})))*24=0,\" \",IF({}=0,IF({}=0,\"0\",{}-{}),IF({}=0,{}-{},({}-{})+({}-{})))*24)",
        cell_ref_1, cell_ref_3, cell_ref_3, cell_ref_4,
        cell_ref_1, cell_ref_1, cell_ref_2,
        cell_ref_1, cell_ref_2, cell_ref_3, cell_ref_4,
        cell_ref_1, cell_ref_3, cell_ref_3, cell_ref_4,
        cell_ref_1, cell_ref_1, cell_ref_2,
        cell_ref_1, cell_ref_2, cell_ref_3, cell_ref_4
    )
}
// ---------------------------------------------------------------------------------//