use serde::{Deserialize, Serialize};
use serde_json;
use std::fs;
use std::path::Path;
use std::path::PathBuf;
use rfd::MessageDialog;
use rfd::MessageLevel;

//--------- Costanti ---------//
const FILE_NAME: &str = "dati_utente.json";
const DIR_NAME: &str = "TS_compiler";
const DEFAULT_NOMINATIVO: &str = "inserire tuo nominativo in dati_utente.json";
const DEFAULT_CLIENTE: &str = "inserire il tuo cliente in dati_utente.json";
const ERROR_DIRECTORY: &str = "Errore nella creazione della directory";
const ERROR_FILE_CREATION: &str = "Errore nella scrittura del file JSON";
const DESERIALIZATION_ERROR: &str = "Errore nella deserializzazione del JSON";
const READ_ERROR: &str = "Errore nella lettura del file JSON";
const DAYS_ERROR: &str = "Errore : Il vettore giorni_sede all'interno del json contiene un numero maggiore di 5 o più di 5 elementi.";
const TITLE_ERROR: &str = "Errore nella valorizzazione dei giorni sede";


//--------- Struct e Impl ---------//
#[derive(Serialize,Deserialize)]
pub struct Utente {
    pub nominativo: String,
    pub cliente: String,
    pub giorni_sede: Vec<u8>,
}
impl Utente {
    pub fn get_utente_from_json() -> Self {
        let file_path = Path::new(DIR_NAME).join(FILE_NAME);
        let contents = fs::read_to_string(&file_path).expect(READ_ERROR);
        serde_json::from_str(&contents).expect(DESERIALIZATION_ERROR)
    }
}
//-------------------------------------------//
pub fn get_file_path (anno:i32,mese:&str) -> String {
        let dir_path = PathBuf::from(DIR_NAME).to_string_lossy().to_string();
        format!("{}/{}",dir_path,handle_file_creation(anno, mese))
    }
//-------------------------------------------//
pub fn handle_file_creation(anno:i32,mese:&str) -> String{
    let dir_path = Path::new(DIR_NAME);
    let file_path = dir_path.join(FILE_NAME);

    // Crea la directory se non esiste
    if !dir_path.exists() {
        fs::create_dir(dir_path).expect(ERROR_DIRECTORY);
    }
    // Crea il file JSON se non esiste
    if !file_path.exists() {
        let utente = Utente {
            nominativo: DEFAULT_NOMINATIVO.to_string(),
            cliente: DEFAULT_CLIENTE.to_string(),
            giorni_sede: vec![
            1,
            2,
            4
            ],
        };
        let json_string = serde_json::to_string_pretty(&utente).unwrap();
        write_in_file(&file_path.to_string_lossy(), &json_string);
    }
    let dati: Utente = Utente::get_utente_from_json();
    check_giorni_sede(&dati.giorni_sede);

    format!("TS_{}_{}_{}.xlsx",dati.nominativo.replace(' ', "_"),mese,anno)
}
//-------------------------------------------//
pub fn check_giorni_sede(dati: &[u8]) {
    if dati.iter().any(|&x| x > 5) || dati.len() > 5 {
        MessageDialog::new()
            .set_title(TITLE_ERROR)
            .set_description(DAYS_ERROR)
            .set_level(MessageLevel::Error)
            .show();
        std::process::exit(1);
    }
}
//-------------------------------------------//
pub fn file_creato() {
        MessageDialog::new()
            .set_title("File creato")
            .set_description("Il file è stato creato con successo.")
            .set_level(MessageLevel::Info)
            .show();
        std::process::exit(1);
}
fn write_in_file(file_path: &str, json_string: &str) {
    match fs::write(file_path, json_string) {
        Ok(_) => (),
        Err(e) => {
        
            MessageDialog::new()
                .set_level(MessageLevel::Error)
                .set_title("Errore scrittura file")
                .set_description(format!("{}:\n{}", ERROR_FILE_CREATION, e))
                .show();

            panic!("{}: {}", ERROR_FILE_CREATION, e);
        }
    }
}
//-------------------------------------------//
pub fn safe_save_workbook(workbook: &mut rust_xlsxwriter::Workbook, nome_file: &str) {
    match workbook.save(nome_file) {
        Ok(_) => (),
        Err(e) => {
        
            MessageDialog::new()
                .set_level(MessageLevel::Error)
                .set_title("Errore scrittura file")
                .set_description(format!("Errore nel salvataggio del file: {}", e))
                .show();

            panic!("{}: {}", ERROR_FILE_CREATION, e);
        }
    }
}
//-------------------------------------------//