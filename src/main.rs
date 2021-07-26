use std::env;
use std::process;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() > 1 && args[1].chars().next() == Some('=') {
        println!("Formula Testing!");
        let lexer = xl::Lexer::new(&args[1]);
        for token in lexer {
            println!("{:?}", token);
        }
        process::exit(0);
    }
    let config = xl::Config::new(&args).unwrap_or_else(|err| {
        match err {
            xl::ConfigError::NeedPathAndTab(_) => {
                eprintln!("Error: {}", err);
                xl::usage();
            },
            xl::ConfigError::NeedTab => {
                eprintln!("Error: {}", err);
                if let Ok(mut wb) = xl::Workbook::open(&args[1]) {
                    eprintln!("The following sheets are available in '{}':", &args[1]);
                    for sheet_name in wb.sheets().by_name() {
                        eprintln!("   {}", sheet_name);
                    }
                } else {
                    eprintln!("(that workbook also does not seem to exist or is not a valid xlsx file)");
                }
                eprintln!("\nSee help by using -h flag.");
            },
            _ => {
                eprintln!("Error: {}", err);
                eprintln!("\nSee help by using -h flag.");
            },
        }
        process::exit(1);
    });
    if let Err(e) = xl::run(config) {
        eprintln!("Runtime error: {}", e);
        process::exit(1);
    }
}
