use std::env;
use std::process;
use xl;

fn main() {
    let args: Vec<String> = env::args().collect();
    let config = xl::Config::new(&args).unwrap_or_else(|err| {
        eprintln!("Error: {}", err);
        process::exit(1);
    });
    if let Err(e) = xl::run(config) {
        eprintln!("Runtime error: {}", e);
        process::exit(1);
    }
}
