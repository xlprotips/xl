
use crate::utils;

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
        RowIter { worksheet_reader: reader }
    }
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

#[derive(Debug)]
pub struct Row<'a>(pub Vec<Cell<'a>>);

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

impl<'a> Iterator for RowIter<'a> {
    type Item = Row<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        let mut buf = Vec::new();
        let reader = &mut self.worksheet_reader.reader;
        let strings = self.worksheet_reader.strings;
        let next_row = {
            let mut row = Vec::new();
            let mut in_cell = false;
            let mut in_value = false;
            let mut c = new_cell();
            loop {
                match reader.read_event(&mut buf) {
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
                    Ok(Event::End(ref e)) if e.name() == b"v" => {
                        in_value = false;
                    },
                    Ok(Event::End(ref e)) if e.name() == b"c" => {
                        row.push(c);
                        c = new_cell();
                        in_cell = false;
                    },
                    // note: because v elements are children of c elements,
                    // need this check to go before the 'in_cell' check
                    Ok(Event::Text(ref e)) if in_value => {
                        let txt = e.unescape_and_decode(&reader).unwrap();
                        if c.cell_type == "s" {
                            println!("STRINGS!");
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
                    Ok(Event::End(ref e)) if e.name() == b"row" => break row,
                    Ok(Event::Eof) => break row,
                    Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
                    _ => (), // There are several other `Event`s we do not consider here
                }
                buf.clear();
            }
        };
        Some(Row(next_row))
    }
}
