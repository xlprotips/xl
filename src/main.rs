use std::env;
use std::process;

fn main() {
    let args: Vec<String> = env::args().collect();
    let config = Config::new(&args).unwrap_or_else(|err| {
        println!("Error: {}", err);
        process::exit(1);
    });
    if let Err(e) = run(config) {
        println!("Runtime error: {}", e);
        process::exit(1);
    }
}

struct Config {
    workbook_path: String,
    tab: String,
}

impl Config {
    fn new(args: &[String]) -> Result<Config, &str> {
        if args.len() < 2 {
            return Err("must provide workbook name and tab name you want to view")
        } else if args.len() < 3 {
            return Err("must also provide which tab you want to view in workbook")
        }
        let workbook_path = args[1].clone();
        let tab = args[2].clone();
        Ok(Config { workbook_path, tab })
    }
}

fn run(config: Config) -> Result<(), String> {
    match sxl::Workbook::new(&config.workbook_path) {
        Ok(mut wb) => {
            let sheets = wb.sheets();
            if let Some(ws) = sheets.get(&*config.tab) {
                for row in ws.rows(&mut wb).take(6) {
                    println!("{}", row);
                }
            } else {
                return Err("that sheet does not exist".to_owned())
            }
            Ok(())
        },
        Err(e) => Err(e)
    }
}
