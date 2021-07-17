//! Rust library to deal with *big* Excel files.
//!
//! This library is intended to help you deal with big Excel files. The library was originally
//! created as a Python library (https://github.com/ktr/sxl) after learning that neither pandas,
//! openpyxl, xlwings, nor win32com had the ability to open large Excel files without loading them
//! completely into memory. This doesn't work when you have *huge* Excel files (especially if you
//! only want to examine a bit of the file - the first 10 rows say). `sxl` (and this library) solve
//! the problem by parsing the SpreadsheetML / XML xlsx files using a streaming parser. So you can
//! see the first ten rows of any tab within any Excel file extremely quickly.
//!
//! Here is a sample of how you might use this library:
//!
//!     use sxl::Workbook;
//!
//!     fn main () {
//!         let mut wb = sxl::Workbook::open("tests/data/Book1.xlsx").unwrap();
//!         let sheets = wb.sheets();
//!         let sheet = sheets.get("Sheet1");
//!     }

mod ws;
mod utils;

use std::collections::HashMap;
use std::convert::TryInto;
use std::fs;
use std::fs::File;
use std::io::BufReader;
use regex::Regex;
use quick_xml::Reader;
use quick_xml::events::Event;
use zip::ZipArchive;
use zip::read::ZipFile;
pub use ws::Worksheet;


const XL_MAX_COL: u16 = 16384;
const XL_MIN_COL: u16 = 1;


#[cfg(test)]
mod tests {
    mod utility_functions {
        use super::super::*;
        #[test]
        fn num_to_letter_w() {
            assert_eq!(num2col(23), Some(String::from("W")));
        }

        #[test]
        fn num_to_letter_aa() {
            assert_eq!(num2col(27), Some(String::from("AA")));
        }

        #[test]
        fn num_to_letter_ab() {
            assert_eq!(num2col(28), Some(String::from("AB")));
        }

        #[test]
        fn num_to_letter_xfd() {
            assert_eq!(num2col(16384), Some(String::from("XFD")));
        }

        #[test]
        fn num_to_letter_xfe() {
            assert_eq!(num2col(16385), None);
        }

        #[test]
        fn num_to_letter_0() {
            assert_eq!(num2col(0), None);
        }

        #[test]
        fn letter_to_num_w() {
            assert_eq!(col2num("W"), Some(23));
        }

        #[test]
        fn letter_to_num_aa() {
            assert_eq!(col2num("AA"), Some(27));
        }

        #[test]
        fn letter_to_num_ab() {
            assert_eq!(col2num("AB"), Some(28));
        }

        #[test]
        fn letter_to_num_xfd() {
            assert_eq!(col2num("XFD"), Some(16384));
        }

        #[test]
        fn letter_to_num_xfe() {
            assert_eq!(col2num("XFE"), None);
        }

        #[test]
        fn letter_to_num_ab_lower() {
            assert_eq!(col2num("ab"), Some(28));
        }

        #[test]
        fn letter_to_num_number() {
            assert_eq!(col2num("12"), None);
        }

        #[test]
        fn letter_to_num_semicolon() {
            assert_eq!(col2num(";"), None);
        }
    }

    mod access {
        use super::super::*;

        #[test]
        fn open_wb() {
            let wb = Workbook::open("tests/data/Book1.xlsx");
            assert!(wb.is_some());
        }

        #[test]
        fn all_sheets() {
            let mut wb = Workbook::open("tests/data/Book1.xlsx").unwrap();
            let num_sheets = wb.sheets().len();
            assert_eq!(num_sheets, 4);
        }

        #[test]
        fn sheet_by_name_exists() {
            let mut wb = Workbook::open("tests/data/Book1.xlsx").unwrap();
            let sheets = wb.sheets();
            assert!(sheets.get("Time").is_some());
        }

        #[test]
        fn sheet_by_num_exists() {
            let mut wb = Workbook::open("tests/data/Book1.xlsx").unwrap();
            let sheets = wb.sheets();
            assert!(sheets.get(1).is_some());
        }

        #[test]
        fn sheet_by_name_not_exists() {
            let mut wb = Workbook::open("tests/data/Book1.xlsx").unwrap();
            let sheets = wb.sheets();
            assert!(!sheets.get("Unknown").is_some());
        }

        #[test]
        fn sheet_by_num_not_exists() {
            let mut wb = Workbook::open("tests/data/Book1.xlsx").unwrap();
            let sheets = wb.sheets();
            assert!(!sheets.get(0).is_some());
        }

        #[test]
        fn correct_sheet_name() {
            let mut wb = Workbook::open("tests/data/Book1.xlsx").unwrap();
            let sheets = wb.sheets();
            assert_eq!(sheets.get("Time").unwrap().name, "Time");
        }
    }
}


/// Return column letter for column number `n`
pub fn num2col(n: u16) -> Option<String> {
    if n > XL_MAX_COL || n < XL_MIN_COL { return None }
    let mut s = String::new();
    let mut n = n;
    while n > 0 {
        let r: u8 = ((n - 1) % 26).try_into().unwrap();
        n = (n - 1) / 26;
        s.push((65 + r) as char)
    }
    Some(s.chars().rev().collect::<String>())
}


/// Return column number for column letter `letter`
pub fn col2num(letter: &str) -> Option<u16> {
    let letter = letter.to_uppercase();
    let re = Regex::new(r"[A-Z]+").unwrap();
    if !re.is_match(&letter) { return None }
    let mut num: u16 = 0;
    for c in letter.chars() {
        num = num * 26 + ((c as u16) - ('A' as u16)) + 1;
    }
    if num > XL_MAX_COL || num < XL_MIN_COL { return None }
    Some(num)
}

#[derive(Debug)]
pub enum DateSystem {
    V1900,
    V1904,
}

#[derive(Debug)]
pub struct Workbook {
    pub path: String,
    xls: ZipArchive<fs::File>,
    pub encoding: String,
    pub date_system: DateSystem,
    pub strings: Vec<String>,
}

#[derive(Debug)]
pub struct SheetMap {
    sheets_by_name: HashMap::<String, usize>,
    sheets_by_num: Vec<Option<Worksheet>>,
}

pub enum Sheet<'a> {
    Name(&'a str),
    Pos(usize),
}

pub trait SheetTrait { fn go(&self) -> Sheet; }

impl SheetTrait for &str {
    fn go(&self) -> Sheet { Sheet::Name(*self) }
}

impl SheetTrait for usize {
    fn go(&self) -> Sheet { Sheet::Pos(*self) }
}

impl SheetMap {
    pub fn get<T: SheetTrait>(&self, sheet: T) -> Option<&Worksheet> {
        let sheet = sheet.go();
        match sheet {
            Sheet::Name(n) => {
                match self.sheets_by_name.get(n) {
                    Some(p) => self.sheets_by_num.get(*p)?.as_ref(),
                    None => None
                }
            },
            Sheet::Pos(n) => self.sheets_by_num.get(n)?.as_ref(),
        }
    }

    pub fn len(&self) -> u8 {
        (self.sheets_by_num.len() - 1) as u8
    }
}

pub struct SheetReader<'a> {
    reader: Reader<BufReader<ZipFile<'a>>>,
    strings: &'a Vec<String>
}

impl Workbook {
    /// xlsx zips contain an xml file that has a mapping of "ids" to "targets." The ids are used
    /// to uniquely identify sheets within the file. The targets have information on where the
    /// sheets can be found within the zip. This function returns a hashmap of id -> target so that
    /// you can quickly determine the name of the sheet xml file within the zip.
    fn rels(&mut self) -> HashMap<String, String> {
        let mut map = HashMap::new();
        match self.xls.by_name("xl/_rels/workbook.xml.rels") {
            Ok(rels) => {
                // Looking for tree structure like:
                //   Relationships
                //     Relationship(id = "abc", target = "def")
                //     Relationship(id = "ghi", target = "lkm")
                //     etc.
                //  Each relationship contains an id that is used to reference
                //  the sheet and a target which tells us where we can find the
                //  sheet in the zip file.
                //
                //  Uncomment the following line to print out a copy of what
                //  the xml looks like (will probably not be too big).
                // let _ = std::io::copy(&mut rels, &mut std::io::stdout());

                let reader = BufReader::new(rels);
                let mut reader = Reader::from_reader(reader);
                reader.trim_text(true);

                let mut buf = Vec::new();
                loop {
                    match reader.read_event(&mut buf) {
                        Ok(Event::Empty(ref e)) => {
                            match e.name() {
                                b"Relationship" => {
                                    let mut id = String::new();
                                    let mut target = String::new();
                                    e.attributes()
                                        .for_each(|a| {
                                            let a = a.unwrap();
                                            if a.key == b"Id" {
                                                id = utils::attr_value(&a);
                                            }
                                            if a.key == b"Target" {
                                                target = utils::attr_value(&a);
                                            }
                                        });
                                    map.insert(id, target);
                                },
                                _ => (),
                            }
                        },
                        Ok(Event::Eof) => break, // exits the loop when reaching end of file
                        Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
                        _ => (), // There are several other `Event`s we do not consider here
                    }
                    buf.clear();
                }

                map
            },
            Err(_) => map
        }
    }

    /// Return hashmap of all sheets (sheet name -> Worksheet)
    pub fn sheets(&mut self) -> SheetMap {
        let rels = self.rels();
        let num_sheets = rels.iter().filter(|(_, v)| v.starts_with("worksheet")).count();
        let mut sheets = SheetMap {
            sheets_by_name: HashMap::new(),
            sheets_by_num: Vec::with_capacity(num_sheets + 1), // never a "0" sheet
        };

        match self.xls.by_name("xl/workbook.xml") {
            Ok(wb) => {
                // let _ = std::io::copy(&mut wb, &mut std::io::stdout());

                let reader = BufReader::new(wb);
                let mut reader = Reader::from_reader(reader);
                reader.trim_text(true);

                let mut buf = Vec::new();
                loop {
                    match reader.read_event(&mut buf) {
                        Ok(Event::Empty(ref e)) if e.name() == b"sheet" => {
                            let mut name = String::new();
                            let mut id = String::new();
                            let mut num = 0;
                            e.attributes()
                                .for_each(|a| {
                                    let a = a.unwrap();
                                    if a.key == b"r:id" {
                                        id = utils::attr_value(&a);
                                    }
                                    if a.key == b"name" {
                                        name = utils::attr_value(&a);
                                    }
                                    if a.key == b"sheetId" {
                                        if let Ok(r) = utils::attr_value(&a).parse() {
                                            num = r;
                                        }
                                    }
                                });
                            sheets.sheets_by_name.insert(name.clone(), num as usize);
                            let target = {
                                let s = rels.get(&id).unwrap();
                                if s.starts_with("/") {
                                    s[1..].to_string()
                                } else {
                                    "xl/".to_owned() + s
                                }
                            };
                            while num as usize >= sheets.sheets_by_num.len() {
                                sheets.sheets_by_num.push(None);
                            }
                            sheets.sheets_by_num[num as usize] = Some(Worksheet::new(id, name, num, target));
                        },
                        Ok(Event::Eof) => break,
                        Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
                        _ => (),
                    }
                    buf.clear();
                }
                sheets
            },
            Err(_) => sheets
        }
    }

    pub fn new(path: &str) -> Option<Self> {
        if !std::path::Path::new(&path).exists() { return None }
        let zip_file = fs::File::open(&path).unwrap();
        if let Ok(mut xls) = zip::ZipArchive::new(zip_file) {
            let strings = strings(&mut xls);
            Some(Workbook {
                path: path.to_string(),
                xls,
                encoding: String::from("utf8"),
                date_system: DateSystem::V1900,
                strings,
            })
        } else {
            return None
        }
    }

    pub fn open(path: &str) -> Option<Self> { Workbook::new(path) }

    pub fn contents(&mut self) {
        for i in 0 .. self.xls.len() {
            let file = self.xls.by_index(i).unwrap();
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

    pub fn sheet_reader<'a>(&'a mut self, zip_target: &str) -> SheetReader<'a> {
        let target = match self.xls.by_name(zip_target) {
            Ok(ws) => ws,
            Err(_) => panic!("Could not find worksheet: {}", zip_target)
        };
        // let _ = std::io::copy(&mut target, &mut std::io::stdout());
        let reader = BufReader::new(target);
        let mut reader = Reader::from_reader(reader);
        reader.trim_text(true);
        SheetReader { reader, strings: &self.strings }
        // SheetReader { reader, get_string: &self.get_string }
    }

}


fn strings(zip_file: &mut ZipArchive<File>) -> Vec<String> {
    let mut strings = Vec::new();
    match zip_file.by_name("xl/sharedStrings.xml") {
        Ok(strings_file) => {
            let reader = BufReader::new(strings_file);
            let mut reader = Reader::from_reader(reader);
            reader.trim_text(true);
            let mut buf = Vec::new();
            loop {
                match reader.read_event(&mut buf) {
                    Ok(Event::Text(ref e)) => strings.push(e.unescape_and_decode(&reader).unwrap()),
                    Ok(Event::Eof) => break, // exits the loop when reaching end of file
                    Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
                    _ => (), // There are several other `Event`s we do not consider here
                }
                buf.clear();
            }
            strings
        },
        Err(_) => strings
    }
}


/*
 * # ISO/IEC 29500:2011 in Part 1, section 18.8.30
STANDARD_STYLES = {
    '0' : 'General',
    '1' : '0',
    '2' : '0.00',
    '3' : '#,##0',
    '4' : '#,##0.00',
    '9' : '0%',
    '10' : '0.00%',
    '11' : '0.00E+00',
    '12' : '# ?/?',
    '13' : '# ??/??',
    '14' : 'mm-dd-yy',
    '15' : 'd-mmm-yy',
    '16' : 'd-mmm',
    '17' : 'mmm-yy',
    '18' : 'h:mm AM/PM',
    '19' : 'h:mm:ss AM/PM',
    '20' : 'h:mm',
    '21' : 'h:mm:ss',
    '22' : 'm/d/yy h:mm',
    '37' : '#,##0 ;(#,##0)',
    '38' : '#,##0 ;[Red](#,##0)',
    '39' : '#,##0.00;(#,##0.00)',
    '40' : '#,##0.00;[Red](#,##0.00)',
    '45' : 'mm:ss',
    '46' : '[h]:mm:ss',
    '47' : 'mmss.0',
    '48' : '##0.0E+0',
    '49' : '@',
}
*/
