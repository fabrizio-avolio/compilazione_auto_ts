use chrono::{Local, Datelike, NaiveDate, Weekday};

const DATA_NON_VALIDA: &str = "DATA NON VALIDA";
#[derive(Debug)]
pub struct GiornoCalendario<'a> {
    pub giorno: u32,
    pub giorno_settimana: &'a str,
}

#[derive(Debug)]
pub struct Calendario<'a> {
    pub anno: i32,
    pub mese: &'a str,
    pub giorni: Vec<GiornoCalendario<'a>>,
}

fn tradurre_giorno_settimana(giorno: Weekday) -> &'static str {
    match giorno {
        Weekday::Mon => "lunedì",
        Weekday::Tue => "martedì",
        Weekday::Wed => "mercoledì",
        Weekday::Thu => "giovedì",
        Weekday::Fri => "venerdì",
        Weekday::Sat => "sabato",
        Weekday::Sun => "domenica",
    }
}
fn tradurre_mese(mese: u32) -> &'static str {
    let nome_mese:&str = match mese {
        1 => "Gennaio",
        2 => "Febbraio",
        3 => "Marzo",
        4 => "Aprile",
        5 => "Maggio",
        6 => "Giugno",
        7 => "Luglio",
        8 => "Agosto",
        9 => "Settembre",
        10 => "Ottobre",
        11 => "Novembre",
        12 => "Dicembre",
        _ => "Mese non valido", // Gestione di input non validi
    };
    nome_mese
}


impl<'a> Calendario<'a> {
    pub fn new() -> Self {
        let today = Local::now().date_naive();
        let anno = today.year();
        let mese1:u32 = today.month();
        let mese = tradurre_mese(mese1);
        
        let first_day = NaiveDate::from_ymd_opt(anno, mese1, 1).expect(DATA_NON_VALIDA);

        let giorni_del_mese = if mese1 == 12 {
            NaiveDate::from_ymd_opt(anno + 1, 1, 1)
                .expect(DATA_NON_VALIDA)
                .signed_duration_since(first_day)
                .num_days() as u32
        } else {
            NaiveDate::from_ymd_opt(anno, mese1 + 1, 1)
                .expect(DATA_NON_VALIDA)
                .signed_duration_since(first_day)
                .num_days() as u32
        };

        let mut giorni = Vec::new();
        for giorno in 1..=giorni_del_mese {
            let giorno_settimana = tradurre_giorno_settimana(NaiveDate::from_ymd_opt(anno, mese1, giorno).expect(DATA_NON_VALIDA).weekday());
            giorni.push(GiornoCalendario { giorno, giorno_settimana });
        }

        Calendario { anno, mese, giorni }
    }
}
