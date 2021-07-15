
use std::fmt;
use quick_xml::events::Event;
// use quick_xml::events::attributes::Attribute;
use crate::{SheetReader, Workbook};

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
}

impl Worksheet {
    pub fn new(id: String, name: String, position: u8, target: String) -> Self {
        Worksheet { name, position, id, target }
    }

    pub fn rows<'a>(&self, workbook: &'a mut Workbook) -> RowIter<'a> {
        let reader = workbook.sheet_reader(&self.target);
        RowIter{ worksheet_reader: reader, count: 0 }
    }
}

#[derive(Debug)]
pub enum ExcelValue {
    Bool(String),
    Date(String),
    Err,
    None,
    Number(f64),
    String(String),
}

impl fmt::Display for ExcelValue {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            ExcelValue::Bool(b) => write!(f, "{}", b),
            ExcelValue::Number(n) => write!(f, "{}", n),
            ExcelValue::String(s) => write!(f, "\"{}\"", s),
            ExcelValue::Err => write!(f, "#NA"),
            ExcelValue::Date(d) => write!(f, "'{}'", d),
            ExcelValue::None => write!(f, "<None>"),
        }
    }
}

#[derive(Debug)]
pub struct Cell {
    pub value: ExcelValue,
    pub formula: String,
}

#[derive(Debug)]
pub struct Row(pub Vec<Cell>);

impl fmt::Display for Row {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        let vec = &self.0;
        write!(f, "[")?;
        /*
        self.0.iter().fold(Ok(()), |result, cell| {
            result.and_then(|_| write!(f, "--> {}; ", cell))
        })
        */
        for (count, v) in vec.iter().enumerate() {
            if count != 0 { write!(f, ", ")?; }
            write!(f, "{}", v)?;
        }
        write!(f, "]")
    }
}

impl fmt::Display for Cell {
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
    count: u32,
}

impl Iterator for RowIter<'_> {
    type Item = Row;

    fn next(&mut self) -> Option<Self::Item> {
        if self.count < 5 {
            self.count += 1;

            let mut buf = Vec::new();
            let next_row = loop {
                let mut row = Vec::new();
                match self.worksheet_reader.read_event(&mut buf) {
                    Ok(Event::Empty(ref e)) => {
                        match e.name() {
                            b"c" => {
                                let c = Cell {
                                    value: ExcelValue::None,
                                    formula: "".to_string(),
                                };
                                row.push(c)
                            },
                            _ => (),
                        }
                    },
                    Ok(Event::End(ref e)) => {
                        match e.name() {
                            b"row" => break row,
                            _ => ()
                        }
                    },
                    Ok(Event::Eof) => {
                        break row
                    },
                    Err(e) => panic!("Error at position {}: {:?}", self.worksheet_reader.buffer_position(), e),
                    _ => (), // There are several other `Event`s we do not consider here
                }
                buf.clear();
            };
            Some(Row(next_row))

        } else {
            None
        }
    }
}
