use std::convert::TryInto;
use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};
use quick_xml::events::attributes::{Attribute, Attributes};
use crate::wb::DateSystem;

const XL_MAX_COL: u16 = 16384;
const XL_MIN_COL: u16 = 1;

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
    let mut num: u16 = 0;
    for c in letter.chars() {
        if c < 'A' || c > 'Z' { return None }
        num = num * 26 + ((c as u16) - ('A' as u16)) + 1;
    }
    if num > XL_MAX_COL || num < XL_MIN_COL { return None }
    Some(num)
}

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
    Date(NaiveDate),
    DateTime(NaiveDateTime),
    Time(NaiveTime),
    Number(i64),
}

///  Return date of "number" based on the date system provided.
///
///  The date system is either the 1904 system or the 1900 system depending on which date system
///  the spreadsheet is using. See <http://bit.ly/2He5HoD> for more information on date systems in
///  Excel.
pub fn excel_number_to_date(number: f64, date_system: &DateSystem) -> DateConversion {
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
        if date.time() == NaiveTime::from_hms(0, 0, 0) {
            DateConversion::Date(date.date())
        } else {
            DateConversion::DateTime(date)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

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
