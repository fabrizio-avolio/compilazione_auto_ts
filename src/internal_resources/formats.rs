use rust_xlsxwriter::{Format, FormatAlign, FormatBorder, Color};
pub fn times_new_roman_bold_italic() -> Format {
    rust_xlsxwriter::Format::new()
        .set_bold()
        .set_italic()
        .set_font_name("Times New Roman")
        .set_font_size(12)
}
pub fn times_new_roman_italic_with_font(font: i32) -> Format {
    rust_xlsxwriter::Format::new()
        .set_italic()
        .set_font_name("Times New Roman")
        .set_font_size(font)
}

pub fn times_new_centered() -> Format {
    rust_xlsxwriter::Format::new()
        .set_align(FormatAlign::Center)
        .set_font_name("Times New Roman")
        .set_border_right(FormatBorder::Thin)
        .set_border_bottom(FormatBorder::Thin)
        .set_border_top(FormatBorder::Thin)
}

pub fn times_new_centered_bolded() -> Format {
    rust_xlsxwriter::Format::new()
        .set_align(FormatAlign::Center)
        .set_bold()
        .set_font_name("Times New Roman")
        .set_border_right(FormatBorder::Medium)
        .set_border_bottom(FormatBorder::Medium)
        .set_border_top(FormatBorder::Medium)
        .set_border_left(FormatBorder::Medium)
}

pub fn times_new_roman_italic() -> Format {
    rust_xlsxwriter::Format::new()
        .set_italic()
        .set_font_name("Times New Roman")
        .set_font_size(12)
}

pub fn times_new_roman_bold_underline() -> Format {
    rust_xlsxwriter::Format::new()
        .set_underline(rust_xlsxwriter::FormatUnderline::Single) // Specifica il tipo di sottolineatura
        .set_font_name("Times New Roman")
        .set_font_size(13)
        .set_bold() // Aggiunge lo stile grassetto
}

pub fn times_new_yellow_centered_bold(font: i32) -> Format {
    rust_xlsxwriter::Format::new()
        .set_align(FormatAlign::Center)
        .set_bold()
        .set_font_size(font)
        .set_font_name("Times New Roman")
        .set_background_color(Color::Yellow)
        .set_border_bottom(FormatBorder::Medium)
        .set_border_top(FormatBorder::Thin)
        .set_border_right(FormatBorder::Thin)
}
pub fn times_new_yellow() -> Format {
    rust_xlsxwriter::Format::new()
        .set_font_name("Times New Roman")
        .set_background_color(Color::Yellow)
        .set_border_bottom(FormatBorder::Thin)
        .set_border_top(FormatBorder::Thin)
        .set_border_right(FormatBorder::Thin)
}

pub fn times_new_green_borded() -> Format {
    rust_xlsxwriter::Format::new()
        .set_font_name("Times New Roman")
        .set_background_color(Color::RGB(0x00FF00))
        .set_border_bottom(FormatBorder::Thin)
        .set_border_top(FormatBorder::Thin)
        .set_border_right(FormatBorder::Thin)
        .set_border_left(FormatBorder::Thin)
        .set_border_bottom(FormatBorder::Thin)
}
pub fn borded_no_right() -> Format {
    rust_xlsxwriter::Format::new()
        .set_font_name("Times New Roman")
        .set_border_bottom(FormatBorder::Thin)
        .set_border_top(FormatBorder::Thin)
        .set_border_right(FormatBorder::Thin)
}
pub fn borded_normal_text(font: i32) -> Format {
    rust_xlsxwriter::Format::new()
        .set_font_name("Times New Roman")
        .set_font_size(font)
        .set_border_bottom(FormatBorder::Thin)
        .set_border_top(FormatBorder::Thin)
        .set_border_right(FormatBorder::Thin)
        .set_border_left(FormatBorder::Thin)
}
pub fn borded_bold_text(font: i32) -> Format {
    rust_xlsxwriter::Format::new()
        .set_font_name("Times New Roman")
        .set_bold()
        .set_font_size(font)
        .set_border_bottom(FormatBorder::Thin)
        .set_border_top(FormatBorder::Thin)
        .set_border_right(FormatBorder::Thin)
        .set_border_left(FormatBorder::Thin)
}

pub fn borded_bold_text_bottom_medium(font: i32) -> Format {
    rust_xlsxwriter::Format::new()
        .set_font_name("Times New Roman")
        .set_bold()
        .set_font_size(font)
        .set_border_bottom(FormatBorder::Medium)
        .set_border_top(FormatBorder::Thin)
        .set_border_right(FormatBorder::Thin)
        .set_border_left(FormatBorder::Thin)
}

pub fn centered_text_bold() -> Format {
    rust_xlsxwriter::Format::new()
        .set_font_name("Times New Roman")
        .set_align(FormatAlign::Center)
        .set_bold()
}

pub fn top_bordered_medium_bold() -> Format {
    rust_xlsxwriter::Format::new()
        .set_font_name("Times New Roman")
        .set_bold()
        .set_font_size(10)
        .set_border_bottom(FormatBorder::Thin)
        .set_border_top(FormatBorder::Medium)
        .set_border_right(FormatBorder::Thin)
        .set_border_left(FormatBorder::Thin)
}
pub fn bordered_no_left() -> Format {
    rust_xlsxwriter::Format::new()
        .set_font_name("Times New Roman")
        .set_font_size(10)
        .set_border_bottom(FormatBorder::Thin)
        .set_border_top(FormatBorder::Thin)
        .set_border_right(FormatBorder::Thin)
}
