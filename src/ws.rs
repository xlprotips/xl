
use std::fmt;
use crate::Workbook;

#[derive(Debug)]
pub struct Worksheet<'a> {
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
    wb: &'a Workbook,
}

impl<'a> Worksheet<'a> {
    pub fn new(wb: &'a Workbook, id: String, name: String, position: u8, target: String) -> Self {
        Worksheet { wb, name, position, id, target }
    }

    pub fn rows(&self) -> RowIter {
        RowIter{ count: 0 }
    }
}

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

pub struct Cell {
    pub value: ExcelValue,
    pub formula: String,
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

pub struct RowIter {
    count: u32,
}

impl Iterator for RowIter {
    type Item = Cell;

    fn next(&mut self) -> Option<Self::Item> {
        if self.count < 5 {
            self.count += 1;
            let v = Cell { value: ExcelValue::None, formula: "".to_string() };
            Some(v)
        } else {
            None
        }
    }
}
