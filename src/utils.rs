use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};
use quick_xml::events::attributes::{Attribute, Attributes};
use crate::DateSystem;

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

pub enum DateConversion {
    DateTime(NaiveDateTime),
    Time(NaiveTime),
    Number(i64),
}

///  Return date of "number" based on the date system provided.
///
///  The date system is either the 1904 system or the 1900 system depending on which date system
///  the spreadsheet is using. See http://bit.ly/2He5HoD for more information on date systems in
///  Excel.
pub fn excel_number_to_date(number: f64, date_system: DateSystem) -> DateConversion {
    let base = match date_system {
        DateSystem::V1900 => {
            // Under the 1900 base system, 1 represents 1/1/1900 (so we start with a base date of
            // 12/31/1899).
            let mut base = NaiveDate::from_ymd(1899, 12, 31).and_hms(0, 0, 0);
            // BUT (!), Excel considers 1900 a leap-year which it is not. As such, it will happily
            // represent 2/29/1900 with the number 60, but we cannot convert that value to a date
            // so we throw an error.
            if number == 60.0 {
                panic!("Bad date in Excel file - 2/29/1900 not valid")
            // Otherwise, if the value is greater than 60 we need to adjust the base date to
            // 12/30/1899 to account for this leap year bug.
            } else if number > 60.0 {
                base = base - Duration::days(1)
            }
            base
        },
        DateSystem::V1904 => {
            // Under the 1904 system, 1 represent 1/2/1904 so we start with a base date of
            // 1/1/1904.
            NaiveDate::from_ymd(1904, 1, 1).and_hms(0, 0, 0)
        }
    };
    let days = number.trunc() as i64;
    if days < -693594 {
        return DateConversion::Number(days)
    }
    let partial_days = number - (days as f64);
    let seconds = (partial_days * 86400000.0).round() as i64;
    let milliseconds = Duration::milliseconds(seconds % 1000);
    let seconds = Duration::seconds(seconds / 1000);
    let date = base + Duration::days(days) + seconds + milliseconds;
    if days == 0 {
        DateConversion::Time(date.time())
    } else {
        DateConversion::DateTime(date)
    }
}
