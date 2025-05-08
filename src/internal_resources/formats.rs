use rust_xlsxwriter::{Format, Workbook, Worksheet, XlsxAlign, XlsxBorder, XlsxColor, XlsxError};
pub fn times_new_roman_bold_italic() -> Format {
    rust_xlsxwriter::Format::new()
    .set_bold()
    .set_italic()
    .set_font_name("Times New Roman")
    .set_font_size(12)
}
pub fn times_new_roman_bold() -> Format {
    rust_xlsxwriter::Format::new()
    .set_bold()
    .set_font_name("Times New Roman")
    .set_font_size(12)
    .set_border_top(XlsxBorder::Medium)
}

pub fn times_new_centered() -> Format {
    rust_xlsxwriter::Format::new()
    .set_align(XlsxAlign::Center)
    .set_font_name("Times New Roman")
    .set_border_right(XlsxBorder::Thin)
    .set_border_bottom(XlsxBorder::Thin)
    .set_border_top(XlsxBorder::Thin)
}
pub fn times_new_roman_italic() -> Format {
    rust_xlsxwriter::Format::new()
    .set_italic()
    .set_font_name("Times New Roman")
    .set_font_size(12)
}

pub fn times_new_roman_bold_underline() -> Format {
    rust_xlsxwriter::Format::new()
        .set_underline(rust_xlsxwriter::XlsxUnderline::Single) // Specifica il tipo di sottolineatura
        .set_font_name("Times New Roman")
        .set_font_size(13)
        .set_bold() // Aggiunge lo stile grassetto
}

pub fn set_border_top() -> Format {
    Format::new().set_border_top(XlsxBorder::Medium)
}

pub fn set_border_bottom() -> Format {
    Format::new().set_border_bottom(XlsxBorder::Thin)
}

pub fn set_border_left() -> Format {
    Format::new().set_border_left(XlsxBorder::Thin)
}

pub fn set_border_right() -> Format {
    Format::new().set_border_right(XlsxBorder::Thin)
}
pub fn times_new_yellow_centered() -> Format {
    rust_xlsxwriter::Format::new()
    .set_align(XlsxAlign::Center)
    .set_font_name("Times New Roman")
    .set_background_color(XlsxColor::Yellow)
}
pub fn times_new_yellow() -> Format {
    rust_xlsxwriter::Format::new()
    .set_font_name("Times New Roman")
    .set_background_color(XlsxColor::Yellow)
    .set_border_bottom(XlsxBorder::Thin)
    .set_border_top(XlsxBorder::Thin)
    .set_border_right(XlsxBorder::Thin)
}