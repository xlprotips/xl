

fn main() {
    std::process::exit(real_main());
}

fn real_main() -> i32 {
    /*
    let args: Vec<_> = std::env::args().collect();
    if args.len() < 2 {
        println!("Usage: {} <filename> [-s <sheetname>]", args[0]);
        return 1;
    }
    let _ = std::path::Path::new(&*args[1]);
    */
    // let path = String::from("sample.xlsx");
    let path = String::from("tests/data/Book1.xlsx");
    match sxl::Workbook::new(&path) {
        Some(mut wb) => {
            let sheets = wb.sheets();
            if let Some(wip) = sheets.get("Sheet1") {
                for row in wip.rows(&mut wb) {
                    println!("{}", row);
                }
            }
        },
        None => println!("Could not open workbook:")
    }

    if 0 == 1 {
        let mut wb = sxl::Workbook::open(&path).unwrap();
        wb.contents();
    }
    return 0;
}
