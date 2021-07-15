use quick_xml::events::attributes::Attribute;

pub fn attr_value(a: &Attribute) -> String {
    String::from_utf8(a.value.to_vec()).unwrap()
}
