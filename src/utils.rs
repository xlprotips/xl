use quick_xml::events::attributes::{Attribute, Attributes};

pub fn attr_value(a: &Attribute) -> String {
    String::from_utf8(a.value.to_vec()).unwrap()
}

pub fn get(attrs: Attributes, which: &[u8]) -> Option<String> {
    for attr in attrs {
        let a = attr.unwrap();
        if a.key == which {
            return Some(attr_value(&a))
        }
    }
    return None
}
