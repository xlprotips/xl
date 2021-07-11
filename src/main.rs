use std::fs;

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
    let path = String::from("sample.xlsx");
    match sxl::Workbook::new(path) {
        Some(mut wb) => {
            let sheets = wb.sheets();
            println!("{:?}", sheets.get(0));
        },
        None => println!("Could not open workbook:")
    }

    let xls = std::path::Path::new("sample.xlsx");
    let tab = fs::File::open(&xls).unwrap();
    let mut archive = zip::ZipArchive::new(tab).unwrap();
    /*

    {
        let wip = archive.by_name("xl/worksheets/sheet3.xml").unwrap();
        let outpath = match wip.enclosed_name() {
            Some(path) => path.to_owned(),
            None => panic!("Could not find tab 'wip'"),
        };
        println!(
            "'wip' tab: \"{}\" ({} bytes)",
            outpath.display(),
            wip.size()
        );
    }
    */

    if 0 == 1 {
        for i in 0..archive.len() {
            let file = archive.by_index(i).unwrap();
            let outpath = match file.enclosed_name() {
                Some(path) => path.to_owned(),
                None => continue,
            };

            if (&*file.name()).ends_with('/') {
                println!("File {}: \"{}\"", i, outpath.display());
            } else {
                println!(
                    "File {}: \"{}\" ({} bytes)",
                    i,
                    outpath.display(),
                    file.size()
                );
            }
        }
    }
    return 0;
}
