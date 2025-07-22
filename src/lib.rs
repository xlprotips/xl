//! This library is intended to help you deal with big Excel files. The library was originally
//! created as a Python library (<https://github.com/ktr/sxl>) after learning that neither pandas,
//! openpyxl, xlwings, nor win32com had the ability to open large Excel files without loading them
//! completely into memory. This doesn't work when you have *huge* Excel files (especially if you
//! only want to examine a bit of the file - the first 10 rows say). `sxl` (and this library) solve
//! the problem by parsing the SpreadsheetML / XML xlsx files using a streaming parser. So you can
//! see the first ten rows of any tab within any Excel file extremely quickly.
//!
//! This particular module provides the plumbing to connect the command-line interface to the xl
//! library code. It parses arguments passed on the command line, determines if we can act on
//! those arguments, and then provides a `Config` object back that can be passed into the `run`
//! function if we can.
//!
//! In order to call `xlcat`, you need to provide a path to a valid workbook and a tab that can be
//! found in that workbook (either by name or by number). You can (optionally) also pass the number
//! of rows you want to see with the `-n` flag (e.g., `-n 10` limits the output to the first ten
//! rows).
//!
//! # Example Usage
//!
//! Here is a sample of how you might use this library:
//!
//!     use xl::Workbook;
//!
//!     fn main () {
//!         let mut wb = xl::Workbook::open("tests/data/Book1.xlsx").unwrap();
//!         let sheets = wb.sheets();
//!         let sheet = sheets.get("Sheet1");
//!     }

mod wb;
mod ws;
mod utils;

use std::fmt;
pub use wb::Workbook;
pub use ws::{Worksheet, ExcelValue};
pub use utils::{col2num, excel_number_to_date, num2col};

#[derive(Debug, Clone, PartialEq)]
pub enum OutputFormat {
    Csv,
    Markdown,
}

enum SheetNameOrNum {
    Name(String),
    Num(usize),
}

pub struct Config {
    /// Which xlsx file should we print?
    workbook_path: String,
    /// Which tab should we print?
    tab: SheetNameOrNum,
    /// How many rows should we print?
    nrows: Option<u32>,
    /// Should we show usage information?
    want_help: bool,
    /// Should we show the current version?
    want_version: bool,
    /// What output format should we use?
    pub output_format: OutputFormat,
}

pub enum ConfigError<'a> {
    NeedPathAndTab(&'a str),
    NeedTab,
    RowsMustBeInt,
    NeedNumRows,
    UnknownFlag(&'a str),
    InvalidFormat(&'a str),
    NeedFormat,
}

impl<'a> fmt::Display for ConfigError<'a> {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            ConfigError::NeedPathAndTab(exe) => write!(f, "need to provide path and tab when running '{}'. See usage below.", exe),
            ConfigError::NeedTab => write!(f, "must also provide which tab you want to view in workbook"),
            ConfigError::RowsMustBeInt => write!(f, "number of rows must be an integer value"),
            ConfigError::NeedNumRows => write!(f, "must provide number of rows when using -n"),
            ConfigError::UnknownFlag(flag) => write!(f, "unknown flag: {}", flag),
            ConfigError::InvalidFormat(fmt) => write!(f, "invalid format '{}'. Valid formats are 'csv' and 'markdown'", fmt),
            ConfigError::NeedFormat => write!(f, "must provide format when using --fmt"),
        }
    }
}

impl Config {
    pub fn new(args: &[String]) -> Result<Config, ConfigError> {
        if args.len() < 2 {
            return Err(ConfigError::NeedPathAndTab(&args[0]))
        } else if args.len() < 3 {
            return match args[1].as_ref() {
                "-h" | "--help" => Ok(Config {
                    workbook_path: "".to_owned(),
                    tab: SheetNameOrNum::Num(0),
                    nrows: None,
                    want_version: false,
                    want_help: true,
                    output_format: OutputFormat::Csv,
                }),
                "-v" | "--version" => Ok(Config {
                    workbook_path: "".to_owned(),
                    tab: SheetNameOrNum::Num(0),
                    nrows: None,
                    want_version: true,
                    want_help: false,
                    output_format: OutputFormat::Csv,
                }),
                _ => Err(ConfigError::NeedTab)
            }
        }
        let workbook_path = args[1].clone();
        let tab = match args[2].parse::<usize>() {
            Ok(num) => SheetNameOrNum::Num(num),
            Err(_) => SheetNameOrNum::Name(args[2].clone())
        };
        let mut config = Config { workbook_path, tab, nrows: None, want_help: false, want_version: false, output_format: OutputFormat::Csv, };
        let mut iter = args[3..].iter();
        while let Some(flag) = iter.next() {
            let flag = &flag[..];
            match flag {
                "-n" => {
                    if let Some(nrows) = iter.next() {
                        if let Ok(nrows) = nrows.parse::<u32>() {
                            config.nrows = Some(nrows)
                        } else {
                            return Err(ConfigError::RowsMustBeInt)
                        }
                    } else {
                        return Err(ConfigError::NeedNumRows)
                    }
                },
                "--fmt" => {
                    if let Some(format) = iter.next() {
                        match format.as_ref() {
                            "csv" => config.output_format = OutputFormat::Csv,
                            "markdown" => config.output_format = OutputFormat::Markdown,
                            _ => return Err(ConfigError::InvalidFormat(format)),
                        }
                    } else {
                        return Err(ConfigError::NeedFormat)
                    }
                },
                _ => return Err(ConfigError::UnknownFlag(flag)),
            }
        }
        Ok(config)
    }
}

pub fn run(config: Config) -> Result<(), String> {
    if config.want_help {
        usage();
        std::process::exit(0);
    }
    if config.want_version {
        version();
        std::process::exit(0);
    }
    match crate::Workbook::new(&config.workbook_path) {
        Ok(mut wb) => {
            let sheets = wb.sheets();
            let sheet = match config.tab {
                SheetNameOrNum::Name(n) => sheets.get(&*n),
                SheetNameOrNum::Num(n) => sheets.get(n),
            };
            if let Some(ws) = sheet {
                let nrows = if let Some(nrows) = config.nrows {
                    nrows as usize
                } else {
                    1048576 // max number of rows in an Excel worksheet
                };
                match config.output_format {
                    OutputFormat::Csv => {
                        for row in ws.rows(&mut wb).take(nrows) {
                            println!("{}", row);
                        }
                    },
                    OutputFormat::Markdown => {
                        // Collect all CSV rows first, then convert to markdown
                        let mut csv_rows: Vec<String> = Vec::new();
                        for row in ws.rows(&mut wb).take(nrows) {
                            let csv_line = format!("{}", row);
                            if !csv_line.trim().is_empty() {
                                csv_rows.push(csv_line);
                            }
                        }
                        
                        // Convert CSV to markdown
                        print_csv_as_markdown(&csv_rows);
                    },
                }
            } else {
                return Err("that sheet does not exist".to_owned())
            }
            Ok(())
        },
        Err(e) => Err(e)
    }
}

pub fn usage() {
    println!(concat!(
        "\n",
        "xlcat 0.1.8\n",
        "Kevin Ryan <ktr@xlpro.tips>\n",
        "\n",
        "xlcat is like cat, but for Excel files (xlsx files to be precise). You simply\n",
        "give it the path of the xlsx and the tab you want to view, and it prints the\n",
        "data in that tab to your screen in a comma-delimited format.\n",
        "\n",
        "You can read about the project at https://xlpro.tips/posts/xlcat. The project\n",
        "page is hosted at https://github.com/xlprotips/xl.\n",
        "\n",
        "USAGE:\n",
        "  xlcat PATH TAB [-n NUM] [--fmt FORMAT] [-h | --help]\n",
        "\n",
        "ARGS:\n",
        "  PATH      Where the xlsx file is located on your filesystem.\n",
        "  TAB       Which tab in the xlsx you want to print to screen.\n",
        "\n",
        "OPTIONS:\n",
        "  -n <NUM>     Limit the number of rows we print to <NUM>.\n",
        "  --fmt FORMAT Output format: 'csv' (default) or 'markdown'.\n",
    ));
}

pub fn version() {
    println!("xlcat 0.1.8");
}

/// Convert CSV rows to markdown table format
fn print_csv_as_markdown(csv_rows: &[String]) {
    if csv_rows.is_empty() {
        return;
    }
    
    let mut rows_data: Vec<Vec<String>> = Vec::new();
    
    // Parse CSV rows
    for csv_row in csv_rows {
        let fields = parse_csv_row(csv_row);
        if !fields.is_empty() && fields.iter().any(|f| !f.trim().is_empty()) {
            rows_data.push(fields);
        }
    }
    
    if rows_data.is_empty() {
        return;
    }
    
    // Find max columns
    let max_cols = rows_data.iter().map(|row| row.len()).max().unwrap_or(0);
    
    // Print header (first row)
    if let Some(header) = rows_data.first() {
        print!("|");
        for i in 0..max_cols {
            let empty_string = String::new();
            let cell = header.get(i).unwrap_or(&empty_string);
            let cleaned = clean_cell_for_markdown(cell);
            print!(" {} |", cleaned);
        }
        println!();
        
        // Print separator row
        print!("|");
        for _ in 0..max_cols {
            print!(" --- |");
        }
        println!();
        
        // Print data rows
        for row in rows_data.iter().skip(1) {
            print!("|");
            for i in 0..max_cols {
                let empty_string = String::new();
                let cell = row.get(i).unwrap_or(&empty_string);
                let cleaned = clean_cell_for_markdown(cell);
                print!(" {} |", cleaned);
            }
            println!();
        }
    }
}

/// Simple CSV parser that handles quoted fields
fn parse_csv_row(csv_row: &str) -> Vec<String> {
    let mut fields = Vec::new();
    let mut current_field = String::new();
    let mut in_quotes = false;
    let mut chars = csv_row.chars().peekable();
    
    while let Some(c) = chars.next() {
        match c {
            '"' => {
                if in_quotes && chars.peek() == Some(&'"') {
                    // Double quote escape
                    current_field.push('"');
                    chars.next(); // consume second quote
                } else {
                    in_quotes = !in_quotes;
                }
            },
            ',' if !in_quotes => {
                fields.push(current_field.trim().to_string());
                current_field.clear();
            },
            _ => current_field.push(c),
        }
    }
    
    // Add the last field
    fields.push(current_field.trim().to_string());
    fields
}

/// Clean cell content for markdown output
fn clean_cell_for_markdown(cell: &str) -> String {
    cell.replace('|', "\\|")
        .replace('\n', " ")
        .replace('\r', " ")
}
