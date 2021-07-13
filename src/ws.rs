
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
}
