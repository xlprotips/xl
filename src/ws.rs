
use crate::utils;

use std::cmp;
use std::fmt;
use std::io::BufReader;
use std::mem;
use chrono::{NaiveDate, NaiveDateTime, NaiveTime};
use zip::read::ZipFile;
use quick_xml::Reader;
use quick_xml::events::Event;
// use quick_xml::events::attributes::Attribute;
use crate::wb::{DateSystem, Workbook};

pub struct SheetReader<'a> {
    reader: Reader<BufReader<ZipFile<'a>>>,
    strings: &'a Vec<String>,
    styles: &'a Vec<String>,
    date_system: &'a DateSystem,
}

impl<'a> SheetReader<'a> {
    pub fn new(
        reader: Reader<BufReader<ZipFile<'a>>>,
        strings: &'a Vec<String>,
        styles: &'a Vec<String>,
        date_system: &'a DateSystem) -> SheetReader<'a> {
        SheetReader { reader, strings, styles, date_system }
    }
}

/// find the number of rows and columns used in a particular worksheet. takes the workbook xlsx
/// location as its first parameter, and the location of the worksheet in question (within the zip)
/// as the second parameter. Returns a tuple of (rows, columns) in the worksheet.
fn used_area(used_area_range: &str) -> (u32, u16) {
    let mut end: isize = -1;
    for (i, c) in used_area_range.chars().enumerate() {
        if c == ':' { end = i as isize; break }
    }
    if end == -1 {
        (0, 0)
    } else {
        let end_range = &used_area_range[end as usize..];
        let mut end = 0;
        // note, the extra '1' (in various spots below) is to deal with the ':' part of the
        // range
        for (i, c) in end_range[1..].chars().enumerate() {
            if !c.is_ascii_alphabetic() {
                end = i + 1;
                break
            }
        }
        let col = utils::col2num(&end_range[1..end]).unwrap();
        let row: u32 = end_range[end..].parse().unwrap();
        (row, col)
    }
}

#[derive(Debug)]
pub struct Worksheet {
    pub name: String,
    pub position: u8,
    relationship_id: String,
    /// location where we can find this worksheet in its xlsx file
    target: String,
    sheet_id: u8,
}

impl Worksheet {
    pub fn new(relationship_id: String, name: String, position: u8, target: String, sheet_id: u8) -> Self {
        Worksheet { name, position, relationship_id, target, sheet_id }
    }

    pub fn rows<'a>(&self, workbook: &'a mut Workbook) -> RowIter<'a> {
        let reader = workbook.sheet_reader(&self.target);
        RowIter {
            worksheet_reader: reader,
            want_row: 1,
            next_row: None,
            num_cols: 0,
            num_rows: 0,
            done_file: false,
        }
    }

}

#[derive(Debug)]
pub enum ExcelValue<'a> {
    Bool(bool),
    Date(NaiveDate),
    DateTime(NaiveDateTime),
    Error(String),
    None,
    Number(f64),
    String(&'a str),
    Time(NaiveTime),
}

impl fmt::Display for ExcelValue<'_> {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            ExcelValue::Bool(b) => write!(f, "{}", b),
            ExcelValue::Date(d) => write!(f, "{}", d),
            ExcelValue::DateTime(d) => write!(f, "{}", d),
            ExcelValue::Error(e) => write!(f, "#{}", e),
            ExcelValue::None => write!(f, ""),
            ExcelValue::Number(n) => write!(f, "{}", n),
            ExcelValue::String(s) => write!(f, "\"{}\"", s),
            ExcelValue::Time(t) => write!(f, "\"{}\"", t),
        }
    }
}

#[derive(Debug)]
pub struct Cell<'a> {
    pub value: ExcelValue<'a>,
    pub formula: String,
    pub reference: String,
    pub style: String,
    pub cell_type: String,
    pub raw_value: String,
}

impl Cell<'_> {
    /// return the row/column coordinates of the current cell
    pub fn coordinates(&self) -> (u16, u32) {
        // let (col, row) = split_cell_reference(&self.reference);
        let (col, row) = {
            let r = &self.reference;
            let mut end = 0;
            for (i, c) in r.chars().enumerate() {
                if !c.is_ascii_alphabetic() {
                    end = i;
                    break
                }
            }
            (&r[..end], &r[end..])
        };
        let col = utils::col2num(col).unwrap();
        let row = row.parse().unwrap();
        (col, row)
    }
}

#[derive(Debug)]
pub struct Row<'a>(pub Vec<Cell<'a>>, pub usize);

impl fmt::Display for Row<'_> {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        let vec = &self.0;
        for (count, v) in vec.iter().enumerate() {
            if count != 0 { write!(f, ",")?; }
            write!(f, "{}", v)?;
        }
        write!(f, "")
    }
}

impl fmt::Display for Cell<'_> {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "{}", self.value)
    }
}

pub struct RowIter<'a> {
    worksheet_reader: SheetReader<'a>,
    want_row: usize,
    next_row: Option<Row<'a>>,
    num_rows: u32,
    num_cols: u16,
    done_file: bool,
}

fn new_cell() -> Cell<'static> {
    Cell {
        value: ExcelValue::None,
        formula: "".to_string(),
        reference: "".to_string(),
        style: "".to_string(),
        cell_type: "".to_string(),
        raw_value: "".to_string(),
    }
}

fn empty_row(num_cols: u16, this_row: usize) -> Option<Row<'static>> {
    let mut row = vec![];
    for n in 0..num_cols {
        let mut c = new_cell();
        c.reference.push_str(&utils::num2col(n + 1).unwrap());
        c.reference.push_str(&this_row.to_string());
        row.push(c);
    }
    Some(Row(row, this_row))
}

impl<'a> Iterator for RowIter<'a> {
    type Item = Row<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        // the xml in the xlsx file will not contain elements for empty rows. So
        // we need to "simulate" the empty rows since the user expects to see
        // them when they iterate over the worksheet.
        if let Some(Row(_, row_num)) = &self.next_row {
            // since we are currently buffering a row, we know we will either return it or a
            // "simulated" (i.e., emtpy) row. So we grab the current row and update the fact that
            // we will soon want a new row. We then figure out if we have the row we want or if we
            // need to keep spitting out empty rows.
            let current_row = self.want_row;
            self.want_row += 1;
            if *row_num == current_row {
                // we finally hit the row we were looking for, so we reset the buffer and return
                // the row that was sitting in it.
                let mut r = None;
                mem::swap(&mut r, &mut self.next_row);
                return r
            } else {
                // otherwise, we must still be sitting behind the row we want. So we return an
                // empty row to simulate the row that exists in the spreadsheet.
                return empty_row(self.num_cols, current_row)
            }
        } else if self.done_file && self.want_row < self.num_rows as usize {
            self.want_row += 1;
            return empty_row(self.num_cols, self.want_row - 1)
        }
        let mut buf = Vec::new();
        let reader = &mut self.worksheet_reader.reader;
        let strings = self.worksheet_reader.strings;
        let styles = self.worksheet_reader.styles;
        let date_system = self.worksheet_reader.date_system;
        let next_row = {
            let mut row: Vec<Cell> = Vec::with_capacity(self.num_cols as usize);
            let mut in_cell = false;
            let mut in_value = false;
            let mut c = new_cell();
            let mut this_row: usize = 0;
            loop {
                match reader.read_event(&mut buf) {
                    /* may be able to get a better estimate for the used area */
                    Ok(Event::Empty(ref e)) if e.name() == b"dimension" => {
                        if let Some(used_area_range) = utils::get(e.attributes(), b"ref") {
                            if used_area_range != "A1" {
                                let (rows, cols) = used_area(&used_area_range);
                                self.num_cols = cols;
                                self.num_rows = rows;
                            }
                        }
                    },
                    /* -- end search for used area */
                    Ok(Event::Start(ref e)) if e.name() == b"row" => {
                        this_row = utils::get(e.attributes(), b"r").unwrap().parse().unwrap();
                    },
                    Ok(Event::Start(ref e)) if e.name() == b"c" => {
                        in_cell = true;
                        e.attributes()
                            .for_each(|a| {
                                let a = a.unwrap();
                                if a.key == b"r" {
                                    c.reference = utils::attr_value(&a);
                                }
                                if a.key == b"t" {
                                    c.cell_type = utils::attr_value(&a);
                                }
                                if a.key == b"s" {
                                    if let Ok(num) = utils::attr_value(&a).parse::<usize>() {
                                        if let Some(style) = styles.get(num) {
                                            c.style = style.to_string();
                                        }
                                    }
                                }
                            });
                    },
                    Ok(Event::Start(ref e)) if e.name() == b"v" => {
                        in_value = true;
                    },
                    // note: because v elements are children of c elements,
                    // need this check to go before the 'in_cell' check
                    Ok(Event::Text(ref e)) if in_value => {
                        c.raw_value = e.unescape_and_decode(&reader).unwrap();
                        c.value = match &c.cell_type[..] {
                            "s" | "str" => {
                                let pos: usize = c.raw_value.parse().unwrap();
                                let s = &strings[pos]; // .to_string()
                                ExcelValue::String(s)
                            },
                            "b" => {
                                if c.raw_value == "0" {
                                    ExcelValue::Bool(false)
                                } else {
                                    ExcelValue::Bool(true)
                                }
                            },
                            "bl" => ExcelValue::None,
                            "e" => ExcelValue::Error(c.raw_value.to_string()),
                            _ if is_date(&c) => {
                                let num = c.raw_value.parse::<f64>().unwrap();
                                match utils::excel_number_to_date(num, date_system) {
                                    utils::DateConversion::Date(date) => ExcelValue::Date(date),
                                    utils::DateConversion::DateTime(date) => ExcelValue::DateTime(date),
                                    utils::DateConversion::Time(time) => ExcelValue::Time(time),
                                    utils::DateConversion::Number(num) => ExcelValue::Number(num as f64),
                                }
                                
                            },
                            _ => ExcelValue::Number(c.raw_value.parse::<f64>().unwrap()),
                        };
                    },
                    Ok(Event::Text(ref e)) if in_cell => {
                        let txt = e.unescape_and_decode(&reader).unwrap();
                        c.formula.push_str(&txt)
                    },
                    Ok(Event::End(ref e)) if e.name() == b"v" => {
                        in_value = false;
                    },
                    Ok(Event::End(ref e)) if e.name() == b"c" => {
                        if let Some(prev) = row.last() {
                            let (mut last_col, _) = prev.coordinates();
                            let (this_col, this_row) = c.coordinates();
                            while this_col > last_col + 1 {
                                let mut cell = new_cell();
                                cell.reference.push_str(&utils::num2col(last_col + 1).unwrap());
                                cell.reference.push_str(&this_row.to_string());
                                row.push(cell);
                                last_col += 1;
                            }
                            row.push(c);
                        } else {
                            let (this_col, this_row) = c.coordinates();
                            for n in 1..this_col {
                                let mut cell = new_cell();
                                cell.reference.push_str(&utils::num2col(n).unwrap());
                                cell.reference.push_str(&this_row.to_string());
                                row.push(cell);
                            }
                            row.push(c);
                        }
                        c = new_cell();
                        in_cell = false;
                    },
                    Ok(Event::End(ref e)) if e.name() == b"row" => {
                        self.num_cols = cmp::max(self.num_cols, row.len() as u16);
                        while row.len() < self.num_cols as usize {
                            let mut cell = new_cell();
                            cell.reference.push_str(&utils::num2col(row.len() as u16 + 1).unwrap());
                            cell.reference.push_str(&this_row.to_string());
                            row.push(cell);
                        }
                        let next_row = Some(Row(row, this_row));
                        if this_row == self.want_row {
                            break next_row
                        } else {
                            self.next_row = next_row;
                            break empty_row(self.num_cols, self.want_row)
                        }
                    },
                    Ok(Event::Eof) => break None,
                    Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
                    _ => (),
                }
                buf.clear();
            }
        };
        self.want_row += 1;
        if next_row.is_none() && self.want_row - 1 < self.num_rows as usize {
            self.done_file = true;
            return empty_row(self.num_cols, self.want_row - 1);
        }
        next_row
    }
}

fn is_date(cell: &Cell) -> bool {
    if cell.style == "d" {
        true
    } else if cell.style.contains("d") && !cell.style.contains("Red") {
        true
    } else if cell.style.contains("m") {
        true
    } else if cell.style.contains("y") {
        true
    } else {
        false
    }
}
