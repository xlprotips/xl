
use crate::utils;

use std::cmp;
use std::fmt;
use std::mem;
use quick_xml::events::Event;
// use quick_xml::events::attributes::Attribute;
use crate::{SheetReader, Workbook};

#[derive(Debug)]
pub struct WorksheetDimensions {
    num_rows: u32,
    num_columns: u16,
}

/// find the number of rows and columns used in a particular worksheet. takes the workbook xlsx
/// location as its first parameter, and the location of the worksheet in question (within the zip)
/// as the second parameter. Returns a tuple of (rows, columns) in the worksheet.
pub fn find_used_area(xlsx: &str, worksheet: &str) -> WorksheetDimensions {
    let mut wb = Workbook::open(xlsx).unwrap();
    let mut reader = wb.sheet_reader(worksheet).reader;
    let mut buf = Vec::new();
    let used_area = loop {
        match reader.read_event(&mut buf) {
            Ok(Event::Empty(ref e)) if e.name() == b"dimension" => {
                let used_area = utils::get(e.attributes(), b"ref").unwrap();
                if used_area != "A1" {
                    break Some(used_area)
                }
            },
            Ok(Event::Start(ref e)) if e.name() == b"sheetData" => {
                break Some("A1:A1".to_string())
            },
            Ok(Event::Eof) => break None,
            Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
            _ => (),
        }
        buf.clear();
    };
    match used_area {
        Some(used_area) => {
            let mut end: isize = -1;
            for (i, c) in used_area.chars().enumerate() {
                if c == ':' { end = i as isize; break }
            }
            if end == -1 {
                WorksheetDimensions { num_rows: 0, num_columns: 0 }
            } else {
                let end_range = &used_area[end as usize..];
                let mut end = 0;
                // note, the extra '1' (in various spots below) is to deal with the ':' part of the
                // range
                for (i, c) in end_range[1..].chars().enumerate() {
                    if !c.is_ascii_alphabetic() {
                        end = i + 1;
                        break
                    }
                }
                let col = crate::col2num(&end_range[1..end]).unwrap();
                let row: u32 = end_range[end..].parse().unwrap();
                WorksheetDimensions {
                    num_rows: row,
                    num_columns: col,
                }
            }
        },
        None => panic!("Could not find used area of worksheet")
    }
}

#[derive(Debug)]
pub struct Worksheet {
    pub name: String,
    pub position: u8,
    id: String,
    // _used_area: 
    // pub row_length: u16,
    // pub num_rows: u32,
    // pub workbook: Workbook,
    // pub name: String,
    // pub position: u8,
    /// location where we can find this worksheet in its xlsx file
    target: String,
    dimensions: WorksheetDimensions,
}

impl Worksheet {
    pub fn new(id: String, name: String, position: u8, target: String, dimensions: WorksheetDimensions) -> Self {
        Worksheet { name, position, id, target, dimensions }
    }

    pub fn rows<'a>(&self, workbook: &'a mut Workbook) -> RowIter<'a> {
        let reader = workbook.sheet_reader(&self.target);
        RowIter {
            worksheet_reader: reader,
            want_row: 1,
            next_row: None,
            num_cols: self.ncols(),
            num_rows: self.nrows(),
            done_file: false,
        }
    }

    /// how many columns are used in this worksheet
    pub fn ncols(&self) -> u16 { self.dimensions.num_columns }

    /// how many rows are used in this worksheet
    pub fn nrows(&self) -> u32 { self.dimensions.num_rows }
}

#[derive(Debug)]
pub enum ExcelValue<'a> {
    Bool(String),
    Date(String),
    Err,
    None,
    Number(f64),
    String(&'a str),
    Other(String),
}

impl fmt::Display for ExcelValue<'_> {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            ExcelValue::Bool(b) => write!(f, "{}", b),
            ExcelValue::Date(d) => write!(f, "'{}'", d),
            ExcelValue::Err => write!(f, "#NA"),
            ExcelValue::None => write!(f, "<None>"),
            ExcelValue::Number(n) => write!(f, "{}", n),
            ExcelValue::Other(s) => write!(f, "\"{}\"", s),
            ExcelValue::String(s) => write!(f, "\"{}\"", s),
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
        let col = crate::col2num(col).unwrap();
        let row = row.parse().unwrap();
        (col, row)
    }
}

#[derive(Debug)]
pub struct Row<'a>(pub Vec<Cell<'a>>, pub usize);

impl fmt::Display for Row<'_> {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        let vec = &self.0;
        write!(f, "[")?;
        for (count, v) in vec.iter().enumerate() {
            if count != 0 { write!(f, ", ")?; }
            write!(f, "{}", v)?;
        }
        write!(f, "]")
    }
}

impl fmt::Display for Cell<'_> {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "{}{}", self.value, if self.formula.starts_with("=") {
            format!(" ({})", self.formula)
        } else {
            if self.formula == "" {
                "".to_string()
            } else {
                format!(" / {}", self.formula)
            }
        })
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
    }
}

fn empty_row(num_cols: u16, this_row: usize) -> Option<Row<'static>> {
    let mut row = vec![];
    for n in 0..num_cols {
        let mut c = new_cell();
        c.reference.push_str(&crate::num2col(n + 1).unwrap());
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
        let next_row = {
            let mut row: Vec<Cell> = Vec::with_capacity(self.num_cols as usize);
            let mut in_cell = false;
            let mut in_value = false;
            let mut c = new_cell();
            let mut this_row: usize = 0;
            loop {
                match reader.read_event(&mut buf) {
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
                                    c.style = utils::attr_value(&a);
                                }
                            });
                    },
                    Ok(Event::Start(ref e)) if e.name() == b"v" => {
                        in_value = true;
                    },
                    // note: because v elements are children of c elements,
                    // need this check to go before the 'in_cell' check
                    Ok(Event::Text(ref e)) if in_value => {
                        let txt = e.unescape_and_decode(&reader).unwrap();
                        if c.cell_type == "s" {
                            let pos: usize = txt.parse().unwrap();
                            let s = &strings[pos]; // .to_string()
                            c.value = ExcelValue::String(s)
                        } else {
                            c.value = ExcelValue::Other(txt)
                        }
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
                                cell.reference.push_str(&crate::num2col(last_col + 1).unwrap());
                                cell.reference.push_str(&this_row.to_string());
                                row.push(cell);
                                last_col += 1;
                            }
                            row.push(c);
                        } else {
                            let (this_col, this_row) = c.coordinates();
                            for n in 1..this_col {
                                let mut cell = new_cell();
                                cell.reference.push_str(&crate::num2col(n).unwrap());
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
                            cell.reference.push_str(&crate::num2col(row.len() as u16 + 1).unwrap());
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
