# Compilazione Auto TS — Automatic Timesheet Generator

![Rust](https://img.shields.io/badge/Rust-2024_edition-orange?logo=rust)
![Excel](https://img.shields.io/badge/output-.xlsx-217346?logo=microsoftexcel&logoColor=white)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)

A lightweight desktop utility written in **Rust** that generates the monthly attendance timesheet
(*Modulo rilevazione presenza del personale*) as a ready-to-submit Excel workbook — in one click.

Filling in the same spreadsheet by hand every month is repetitive and error-prone: copying dates,
working out weekdays, greying out weekends, remembering public holidays, typing the same working hours
and fixing broken formulas. This tool does all of it automatically, so the only thing left to do is
review the file and send it.

---

## Features

- **Zero-input monthly generation** — detects the current month and year and builds one row per calendar day,
  with the Italian weekday name.
- **Pre-filled working hours** — every working day is filled with the standard schedule
  (`08:30–13:00` / `14:00–17:30`).
- **Live Excel formulas** — daily hours, column totals (overtime, holidays, leave, sick days…) and the monthly
  grand total are written as real formulas, so the sheet stays correct after manual edits.
- **Weekend and public holiday detection** — Saturdays, Sundays and all Italian national holidays are
  highlighted and labelled automatically.
- **Easter computed algorithmically** — the movable Easter date is calculated for any year with the
  Anonymous Gregorian (Meeus/Jones/Butcher) algorithm.
- **On-site / remote tracking** — the *Sede* column is set to `SI`/`NO` based on the office days configured
  by the user.
- **Persistent user profile** — name, client and office days are stored in a small JSON file, created
  automatically on first run.
- **Native dialogs** — success and error messages are shown through native OS message boxes, so the tool can
  be launched with a double click, without a terminal.
- **Print-ready layout** — fonts, borders, column widths and colours reproduce the official company template.

## How it works

```text
┌────────────────┐   ┌─────────────────────┐   ┌──────────────────────┐   ┌──────────────────┐
│   Calendario   │──▶│  dati_utente.json   │──▶│   Worksheet builder  │──▶│  TS_<name>_<…>   │
│ current month, │   │ name, client,       │   │ header, daily rows,  │   │      .xlsx       │
│ days, weekdays │   │ office days         │   │ holidays, formulas   │   │                  │
└────────────────┘   └─────────────────────┘   └──────────────────────┘   └──────────────────┘
```

1. `calendario` builds the list of days of the current month with their weekday.
2. `file_handling` loads (or creates) the user profile and validates it.
3. `utility_worksheet` writes the header, one row per day, holiday labels and all the formulas,
   using the styles defined in `formats`.
4. The workbook is saved to `TS_compiler/` and a confirmation dialog is shown.

## Getting started

### Prerequisites

- [Rust](https://www.rust-lang.org/tools/install) **1.85 or newer** (the crate uses the 2024 edition)

### Build

```bash
git clone https://github.com/Royalexperience/compilazione_auto_ts.git
cd compilazione_auto_ts
cargo build --release
```

The executable is created in `target/release/`.

### Run

```bash
cargo run --release
```

or simply double-click the compiled executable.

On the **first run** the tool creates the `TS_compiler/` folder in the current working directory, containing a
`dati_utente.json` template. Fill it in with your details and run the tool again.

## Configuration

`TS_compiler/dati_utente.json`

```json
{
  "nominativo": "Mario Rossi",
  "cliente": "ACME S.p.A.",
  "giorni_sede": [1, 2, 4]
}
```

| Field          | Description                                                                       |
| -------------- | --------------------------------------------------------------------------------- |
| `nominativo`   | Full name of the employee, printed on the sheet and used in the file name.        |
| `cliente`      | Client the hours are reported to.                                                 |
| `giorni_sede`  | Weekdays spent at the office: `1` = Monday … `5` = Friday. Other days are marked as remote. |

`giorni_sede` is validated at startup: it may contain at most five values, each between `1` and `5`.

## Output

The workbook is saved as:

```text
TS_compiler/TS_<Nominativo>_<Mese>_<Anno>.xlsx     e.g. TS_Mario_Rossi_Maggio_2025.xlsx
```

| Column            | Content                                                         |
| ----------------- | --------------------------------------------------------------- |
| **GG**            | Day of the month                                                |
| Weekday           | Italian weekday name                                            |
| **Ent. / Usc.**   | Morning and afternoon clock-in / clock-out times                |
| **Tot. gg**       | Hours worked in the day (formula)                               |
| **Straor. … Malat** | Overtime, holidays, leave, public holidays, time off, sick leave |
| **Sede**          | `SI` if the day is an office day, `NO` otherwise                |
| **N O T E**       | Notes — automatically filled with the holiday name when relevant |

The last rows contain per-column totals and the **total hours of attendance in the month**.

## Project structure

```text
compilazione_auto_ts/
├── Cargo.toml
└── src/
    ├── main.rs                        # Entry point: orchestrates the generation steps
    └── internal_resources/
        ├── calendario.rs              # Month/day model and Italian localisation
        ├── file_handling.rs           # User profile (JSON), validation, saving, dialogs
        ├── formats.rs                 # Reusable cell styles
        └── utility_worksheet.rs       # Sheet layout, holidays, Easter, Excel formulas
```

## Tech stack

| Crate                                                     | Purpose                              |
| --------------------------------------------------------- | ------------------------------------ |
| [`rust_xlsxwriter`](https://crates.io/crates/rust_xlsxwriter) | Excel workbook generation          |
| [`chrono`](https://crates.io/crates/chrono)               | Date and weekday handling            |
| [`serde`](https://crates.io/crates/serde) / [`serde_json`](https://crates.io/crates/serde_json) | User profile (de)serialisation |
| [`rfd`](https://crates.io/crates/rfd)                     | Native message dialogs               |

## Roadmap

- [ ] Choose the month to generate (e.g. via command-line argument) instead of always using the current one
- [ ] Configurable working hours in `dati_utente.json`
- [ ] Easter Monday (*Pasquetta*) and optional local patron-saint holidays
- [ ] Unit tests for the calendar and Easter calculation

## Author

**Fabrizio Avolio** — [GitHub](https://github.com/Royalexperience)
