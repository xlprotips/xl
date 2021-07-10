
use std::fs;

fn main() {
    std::process::exit(real_main());
}

fn real_main() -> i32 {
    let args: Vec<_> = std::env::args().collect();
    if args.len() < 2 {
        println!("Usage: {} <filename> [-s <sheetname>]", args[0]);
        return 1;
    }
    let _ = std::path::Path::new(&*args[1]);
    let fname = std::path::Path::new("sample0.xlsx");
    let file = fs::File::open(&fname).unwrap();
    // let _ = sxl::Worksheet::new(String::from("Test"));

    let mut archive = zip::ZipArchive::new(file).unwrap();

    {
        let file = archive.by_name("wip").unwrap();
        let outpath = match file.enclosed_name() {
            Some(path) => path.to_owned(),
            None => panic!("Could not find tab 'wip'"),
        };
        println!(
            "'wip' tab: \"{}\" ({} bytes)",
            outpath.display(),
            file.size()
        );
    }

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
    return 0;
}
