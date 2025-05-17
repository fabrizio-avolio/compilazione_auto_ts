use std::collections::HashMap;

use rust_xlsxwriter::{Format, FormatAlign, FormatBorder, FormatUnderline, Worksheet};

use super::calendario::{Calendario, GiornoCalendario};
use super::{file_handling, formats};

//-------CONST-------//
const SIZE_ERROR: &str = "Errore impostando la larghezza: ";
const WRITING_ERROR: &str = "Errore nella scrittura su cella: ";
#[derive(Debug, PartialEq)]
struct Festa {
    nome: &'static str,
    descrizione: &'static str,
}

const GIORNI_ROSSI_FISSI: [Festa; 10] = [
    Festa {
        nome: "1 Gennaio",
        descrizione: "Capodanno",
    },
    Festa {
        nome: "6 Gennaio",
        descrizione: "Epifania",
    },
    Festa {
        nome: "25 Aprile",
        descrizione: "Festa della Liberazione",
    },
    Festa {
        nome: "1 Maggio",
        descrizione: "Festa dei Lavoratori",
    },
    Festa {
        nome: "2 Giugno",
        descrizione: "Festa della Repubblica",
    },
    Festa {
        nome: "15 Agosto",
        descrizione: "Ferragosto",
    },
    Festa {
        nome: "1 Novembre",
        descrizione: "Ognissanti",
    },
    Festa {
        nome: "8 Dicembre",
        descrizione: "Immacolata Concezione",
    },
    Festa {
        nome: "25 Dicembre",
        descrizione: "Natale",
    },
    Festa {
        nome: "26 Dicembre",
        descrizione: "Santo Stefano",
    },
];

pub struct UtilityDays {
    giorno_pasqua: String,
    giorni_settimana_for_check: HashMap<&'static str, u8>,
}

impl UtilityDays {
    fn get_utility_days(calendario: &Calendario) -> Self {
        let giorno_pasqua = calcola_pasqua(&calendario.anno);

        let giorni_settimana_for_check: HashMap<&'static str, u8> = [
            ("lunedì", 1),
            ("martedì", 2),
            ("mercoledì", 3),
            ("giovedì", 4),
            ("venerdì", 5),
            ("sabato", 6),
            ("domenica", 7),
        ]
        .iter()
        .cloned()
        .collect();

        UtilityDays {
            giorno_pasqua,
            giorni_settimana_for_check,
        }
    }
}
// ---------------------------------------------------------------------------------//
pub(crate) fn costruisci_gg_nome_gg(worksheet: &mut Worksheet, calendario: &Calendario) {
    let mut i: u32 = 10;
    
    for giorno in &calendario.giorni {
        safe_writing_strings_in_cells(
            worksheet,
            i,
            0,
            &giorno.giorno.to_string(),
            &formats::times_new_centered(),
        );
        colored_yellow_if_checked(
            worksheet,
            i,
            1,
            giorno,
            calendario.mese,
            &UtilityDays::get_utility_days(calendario),
            file_handling::Utente::get_utente_from_json().giorni_sede,
        );
        i += 1;
    }
    let formule: [&str; 7] = [
        "=SUM(G11:G{})",
        "=SUM(H11:H{})",
        "=SUM(I11:I{})",
        "=SUM(J11:J{})",
        "=SUM(K11:K{})",
        "=SUM(L11:L{})",
        "=SUM(M11:M{})",
    ];
    for (k,formula) in formule.iter().enumerate() {
        let new_formula = formula.replace("{}", &i.to_string());
        match k {
            0 => {
                let _ = worksheet.write_formula_with_format(
                    i,
                    (6 + k).try_into().unwrap(),
                    rust_xlsxwriter::Formula::new(&new_formula),
                    &formats::times_new_green_borded()
                        .set_bold()
                        .set_num_format("0.00"),
                );
            }
            _ => {
                let _ = worksheet.write_formula_with_format(
                    i,
                    (6 + k).try_into().unwrap(),
                    rust_xlsxwriter::Formula::new(&new_formula),
                    &formats::borded_bold_text(10).set_num_format("0.00"),
                );
            }
        }
    }
    let somme: [&str; 8] = ["G{}", "H{}", "I{}", "J{}", "K{}", "L{}", "M{}", "N{}"];
    let mut somma_tot: String = String::new();
    for cella in somme {
        somma_tot.push('+');
        let updated_cella = cella.replace("{}", &(i + 1).to_string());
        somma_tot.push_str(&updated_cella);
    }
    somma_tot.insert(0, '=');
    safe_method_to_modify_cell_height(worksheet, i + 1, 9.0);
    safe_writing_strings_in_cells(
        worksheet,
        i + 2,
        0,
        "Totale ore di presenza nel mese :",
        &formats::times_new_roman_italic_with_font(13),
    );
    let _ = worksheet.write_formula_with_format(
        i + 2,
        5,
        rust_xlsxwriter::Formula::new(&somma_tot),
        &formats::times_new_centered_bolded().set_num_format("0.00"),
    );
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
fn safe_writing_strings_in_cells(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    string: &str,
    format: &Format,
) {
    match worksheet.write_string_with_format(row, col, string, format) {
        Ok(_) => (),
        Err(e) => eprintln!("{} {}", WRITING_ERROR, e),
    }
}
// ---------------------------------------------------------------------------------//
fn safe_writing_blank_in_cells(worksheet: &mut Worksheet, row: u32, col: u16, format: &Format) {
    match worksheet.write_blank(row, col, format) {
        Ok(_) => (),
        Err(e) => eprintln!("{} {}", WRITING_ERROR, e),
    }
}
// ---------------------------------------------------------------------------------//
pub fn build_static_strings_in_excel(worksheet: &mut Worksheet) {
    safe_writing_strings_in_cells(
        worksheet,
        1,
        3,
        "Modulo rilevazione presenza del personale",
        &formats::times_new_roman_bold_italic(),
    );

    safe_writing_strings_in_cells(
        worksheet,
        4,
        3,
        "Cliente:",
        &formats::times_new_roman_italic(),
    );

    safe_writing_strings_in_cells(
        worksheet,
        5,
        3,
        "Nominativo",
        &formats::times_new_roman_italic(),
    );

    safe_writing_strings_in_cells(
        worksheet,
        1,
        12,
        "Mese:",
        &formats::times_new_roman_italic(),
    );

    for i in 0..15 {
        match i {
            2 => safe_writing_strings_in_cells(
                worksheet,
                7,
                i,
                "         Orario Entrata-Uscita",
                &formats::top_bordered_medium_bold(),
            ),
            14 => safe_writing_blank_in_cells(
                worksheet,
                7,
                i,
                &rust_xlsxwriter::Format::new()
                    .set_border_top(FormatBorder::Medium)
                    .set_border_right(FormatBorder::Medium),
            ),
            _ => safe_writing_blank_in_cells(worksheet, 7, i, &formats::top_bordered_medium_bold()),
        }
    }

    for i in 0..14 {
        safe_writing_blank_in_cells(
            worksheet,
            9,
            i,
            &Format::new().set_border_bottom(FormatBorder::Medium),
        );
    }

    let intestazioni: [&str; 15] = [
        "",
        "",
        "Mattina",
        "",
        "Pomeriggio",
        "",
        "Ore",
        "Straor.",
        "Ferie",
        "Perm",
        "Fest.",
        "Recupero",
        "Malat",
        "Sede",
        "N O T E",
    ];
    for (i, intestazione) in intestazioni.iter().enumerate() {
        match i {
            2 => {
                safe_writing_strings_in_cells(
                    worksheet,
                    8,
                    i as u16,
                    intestazione,
                    &formats::centered_text_bold()
                        .set_border_left(FormatBorder::Thin)
                        .set_font_size(10),
                );
            }

            3 => {
                safe_writing_strings_in_cells(
                    worksheet,
                    8,
                    i as u16,
                    intestazione,
                    &rust_xlsxwriter::Format::new()
                        .set_border_bottom(FormatBorder::Thin)
                        .set_border_top(FormatBorder::Thin),
                );
            }
            5 => {
                safe_writing_strings_in_cells(
                    worksheet,
                    8,
                    i as u16,
                    intestazione,
                    &rust_xlsxwriter::Format::new()
                        .set_font_name("Times New Roman")
                        .set_border_bottom(FormatBorder::Thin)
                        .set_border_top(FormatBorder::Thin)
                        .set_border_left(FormatBorder::Thin)
                        .set_bold()
                        .set_align(FormatAlign::Center),
                );
            }
            14 => {
                safe_writing_strings_in_cells(
                    worksheet,
                    8,
                    i as u16,
                    intestazione,
                    &formats::centered_text_bold().set_border_right(FormatBorder::Medium),
                );
            }
            _ => {
                safe_writing_strings_in_cells(
                    worksheet,
                    8,
                    i as u16,
                    intestazione,
                    &formats::borded_bold_text(10),
                );
            }
        }
    }
}
// ---------------------------------------------------------------------------------//
pub fn colored_yellow_if_checked(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    giorno: &GiornoCalendario,
    mese: &str,
    utility_days: &UtilityDays,
    giorni_sede: Vec<u8>,
) {
    let giorno_settimana = giorno.giorno_settimana;
    let giorno = giorno.giorno;
    let giorno_pasqua = &utility_days.giorno_pasqua;
    let giorni_settimana_for_check = &utility_days.giorni_settimana_for_check;
    let red_day = format!("{} {}", giorno, mese);
    if giorno_settimana == "sabato"
        || giorno_settimana == "domenica"
        || GIORNI_ROSSI_FISSI
            .iter()
            .any(|festa| festa.nome == red_day.as_str())
        || red_day == *giorno_pasqua
    {
        safe_writing_strings_in_cells(
            worksheet,
            row,
            col,
            giorno_settimana,
            &formats::times_new_yellow(),
        );
        for k in 2..15 {
            let mut format = formats::times_new_yellow();

            if k == 14 {
                if giorno_settimana == "sabato" || giorno_settimana == "domenica" {
                    format = format.set_border_right(FormatBorder::Medium);
                    safe_writing_blank_in_cells(worksheet, row, k, &format);
                } else if red_day == *giorno_pasqua {
                    format = format
                        .set_border_right(FormatBorder::Medium)
                        .set_align(FormatAlign::Center)
                        .set_bold();
                    safe_writing_strings_in_cells(
                        worksheet,
                        row,
                        k,
                        "Pasqua di Resurrezione",
                        &format,
                    );
                } else if let Some(festa) = GIORNI_ROSSI_FISSI
                    .iter()
                    .find(|festa| festa.nome == red_day.as_str())
                {
                    format = format
                        .set_border_right(FormatBorder::Medium)
                        .set_align(FormatAlign::Center)
                        .set_bold();
                    safe_writing_strings_in_cells(worksheet, row, k, festa.descrizione, &format);
                } else {
                    safe_writing_blank_in_cells(worksheet, row, k, &format);
                }
            } else {
                safe_writing_blank_in_cells(worksheet, row, k, &format);
            }
        }
    } else {
        safe_writing_strings_in_cells(
            worksheet,
            row,
            col,
            giorno_settimana,
            &rust_xlsxwriter::Format::new()
                .set_border_bottom(FormatBorder::Thin)
                .set_border_top(FormatBorder::Thin)
                .set_border_right(FormatBorder::Thin),
        );
        safe_writing_strings_in_cells(
            worksheet,
            row,
            col + 1,
            "08:30",
            &formats::borded_normal_text(10).set_align(FormatAlign::Center),
        );
        safe_writing_strings_in_cells(
            worksheet,
            row,
            col + 2,
            "13:00",
            &formats::borded_normal_text(10).set_align(FormatAlign::Center),
        );
        safe_writing_strings_in_cells(
            worksheet,
            row,
            col + 3,
            "14:00",
            &formats::borded_normal_text(10).set_align(FormatAlign::Center),
        );
        safe_writing_strings_in_cells(
            worksheet,
            row,
            col + 4,
            "17:30",
            &formats::borded_normal_text(10).set_align(FormatAlign::Center),
        );
        let formula = generate_formula(row, col + 5);
        let formula_converted = rust_xlsxwriter::Formula::new(formula);
        let _ = worksheet.write_formula_with_format(
            row,
            col + 5,
            formula_converted,
            &formats::borded_bold_text(10).set_num_format("0.00"),
        );
        for k in 7..15 {
            safe_writing_blank_in_cells(worksheet, row, k, &formats::borded_normal_text(10));
            if k == 14 {
                safe_writing_blank_in_cells(
                    worksheet,
                    row,
                    k,
                    &formats::borded_no_right().set_border_right(FormatBorder::Medium),
                )
            }
            if k == 13 {
                if let Some(val) = giorni_settimana_for_check.get(giorno_settimana) {
                    if giorni_sede.contains(val) {
                        safe_writing_strings_in_cells(
                            worksheet,
                            row,
                            k,
                            "SI",
                            &formats::borded_bold_text(10).set_align(FormatAlign::Center),
                        );
                    } else {
                        safe_writing_strings_in_cells(
                            worksheet,
                            row,
                            k,
                            "NO",
                            &formats::borded_bold_text(10).set_align(FormatAlign::Center),
                        );
                    }
                }
            }
        }
    }
}
// ---------------------------------------------------------------------------------//
pub fn build_top_lines(worksheet: &mut Worksheet) {
    let top_intestazioni: [&str; 15] = [
        "GG", "", "Ent.", "Usc.", "Ent.", "Usc.", "Tot. gg", "", "", "", "", "", "", "", "",
    ];
    for (l, top_intestazione) in top_intestazioni.iter().enumerate() {
        match l {
            0 => {
                safe_writing_strings_in_cells(
                    worksheet,
                    9,
                    l as u16,
                    top_intestazione,
                    &formats::times_new_yellow_centered_bold(8),
                );
            }
            1 => {
                safe_writing_strings_in_cells(
                    worksheet,
                    9,
                    l as u16,
                    top_intestazione,
                    &formats::borded_bold_text_bottom_medium(10),
                );
            }
            14 => {
                let format = rust_xlsxwriter::Format::new()
                    .set_font_name("Times New Roman")
                    .set_bold()
                    .set_font_size(10)
                    .set_border_bottom(FormatBorder::Medium)
                    .set_border_left(FormatBorder::Thin)
                    .set_border_right(FormatBorder::Medium);
                safe_writing_strings_in_cells(worksheet, 9, l as u16, top_intestazione, &format);
            }
            _ => {
                safe_writing_strings_in_cells(
                    worksheet,
                    9,
                    l as u16,
                    top_intestazione,
                    &formats::borded_bold_text_bottom_medium(10),
                );
            }
        }
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
        cell_ref_1,
        cell_ref_3,
        cell_ref_3,
        cell_ref_4,
        cell_ref_1,
        cell_ref_1,
        cell_ref_2,
        cell_ref_1,
        cell_ref_2,
        cell_ref_3,
        cell_ref_4,
        cell_ref_1,
        cell_ref_3,
        cell_ref_3,
        cell_ref_4,
        cell_ref_1,
        cell_ref_1,
        cell_ref_2,
        cell_ref_1,
        cell_ref_2,
        cell_ref_3,
        cell_ref_4
    )
}
// ---------------------------------------------------------------------------------//
fn calcola_pasqua(anno: &i32) -> String {
    let a = anno % 19;
    let b = anno / 100;
    let c = anno % 100;
    let d = b / 4;
    let e = b % 4;
    let f = (b + 8) / 25;
    let g = (b - f + 1) / 3;
    let h = (19 * a + b - d - g + 15) % 30;
    let i = c / 4;
    let k = c % 4;
    let l = (32 + 2 * e + 2 * i - h - k) % 7;
    let m = (a + 11 * h + 22 * l) / 451;
    let mese = (h + l - 7 * m + 114) / 31;
    let giorno = ((h + l - 7 * m + 114) % 31) + 1;

    let nomi_mesi = [
        "Gennaio",
        "Febbraio",
        "Marzo",
        "Aprile",
        "Maggio",
        "Giugno",
        "Luglio",
        "Agosto",
        "Settembre",
        "Ottobre",
        "Novembre",
        "Dicembre",
    ];

    format!("{} {}", giorno, nomi_mesi[(mese - 1) as usize])
}
//-------------------------------------------//
pub fn build_month_and_year(worksheet: &mut Worksheet, calendario: &Calendario) {
    safe_writing_strings_in_cells(
        worksheet,
        1,
        14,
        &format!("{} {}", calendario.mese, calendario.anno),
        &formats::times_new_roman_bold_underline(),
    );
}
//-------------------------------------------//
pub fn write_nominativo_and_cliente(worksheet: &mut Worksheet) {
    safe_writing_strings_in_cells(
        worksheet,
        4,
        5,
        &file_handling::Utente::get_utente_from_json().cliente,
        &rust_xlsxwriter::Format::new()
            .set_underline(rust_xlsxwriter::FormatUnderline::Single)
            .set_font_name("Times New Roman")
            .set_font_size(12)
            .set_bold()
            .set_align(FormatAlign::Bottom)
            .set_underline(FormatUnderline::Single),
    );

    safe_writing_strings_in_cells(
        worksheet,
        5,
        5,
        &file_handling::Utente::get_utente_from_json().nominativo,
        &formats::times_new_roman_italic()
            .set_align(FormatAlign::Bottom)
            .set_underline(FormatUnderline::Single),
    );
}
// ---------------------------------------------------------------------------------//



