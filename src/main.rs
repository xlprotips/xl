use std::env;
use std::process;
use xl;

fn main() {
    let args: Vec<String> = env::args().collect();
    let config = xl::Config::new(&args).unwrap_or_else(|err| {
        match err {
            xl::ConfigError::NeedTab => {
                eprintln!("Error: {}", err);
                if let Ok(mut wb) = xl::Workbook::open(&args[1]) {
                    eprintln!("The following sheets are available in '{}':", &args[1]);
                    for sheet_name in wb.sheets().by_name() {
                        eprintln!("   {}", sheet_name);
                    }
                }
            },
            _ => eprintln!("Error: {}", err),
        }
        process::exit(1);
    });
    if let Err(e) = xl::run(config) {
        eprintln!("Runtime error: {}", e);
        process::exit(1);
    }
}
