use std::convert::TryInto;
use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime, Timelike};
use quick_xml::events::attributes::{Attribute, Attributes};
use crate::wb::DateSystem;

const XL_MAX_COL: u16 = 16384;
const XL_MIN_COL: u16 = 1;


/// Return column letter for column number `n`
pub fn num2col(n: u16) -> Option<String> {
    if !(XL_MIN_COL..=XL_MAX_COL).contains(&n) { return None }
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
        if !('A'..='Z').contains(&c) { return None }
        num = num * 26 + ((c as u16) - ('A' as u16)) + 1;
    }
    if !(XL_MIN_COL..=XL_MAX_COL).contains(&num) { return None }
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
    None
}

///  Return date of "number" based on the date system provided.
///
///  The date system is either the 1904 system or the 1900 system depending on which date system
///  the spreadsheet is using. See <http://bit.ly/2He5HoD> for more information on date systems in
///  Excel.
///
///  Some numbers that Excel provides may not properly convert into a date. In such circumstances,
///  we return the representative number of days before the base date that the number represents.
pub fn excel_number_to_date(number: f64, date_system: &DateSystem) -> Result<NaiveDateTime, i64> {
    let base = match date_system {
        DateSystem::V1900 => {
            // Under the 1900 base system, 1 represents 1/1/1900 (so we start with a base date of
            // 12/31/1899).
            let mut base = date_system.base();
            // BUT (!), Excel considers 1900 a leap-year which it is not. As such, it will happily
            // represent 2/29/1900 with the number 60, but we cannot convert that value to a date
            // so we throw an error.
            if (number - 60.0).abs() < 0.0001 {
                panic!("Bad date in Excel file - 2/29/1900 not valid")
            // Otherwise, if the value is greater than 60 we need to adjust the base date to
            // 12/30/1899 to account for this leap year bug.
            } else if number > 60.0 {
                base -= Duration::days(1)
            }
            base
        },
        DateSystem::V1904 => {
            // Under the 1904 system, 1 represent 1/2/1904 so we start with a base date of
            // 1/1/1904.
            date_system.base()
        }
    };
    let days = number.trunc() as i64;
    if days < -693594 {
        return Err(days)
    }
    let partial_days = number - (days as f64);
    let seconds = (partial_days * 86400000.0).round() as i64;
    let milliseconds = Duration::milliseconds(seconds % 1000);
    let seconds = Duration::seconds(seconds / 1000);
    let date = base + Duration::days(days) + seconds + milliseconds;
    Ok(date)
}

pub trait ToDateTime {
    fn to_datetime(&self) -> NaiveDateTime;
}

impl ToDateTime for &NaiveDateTime {
    fn to_datetime(&self) -> NaiveDateTime { **self }
}

impl ToDateTime for &NaiveDate {
    fn to_datetime(&self) -> NaiveDateTime { self.and_hms(0, 0, 0) }
}

impl ToDateTime for &NaiveTime {
    fn to_datetime(&self) -> NaiveDateTime {
        let hour = self.hour();
        let min = self.minute();
        let sec = self.second();
        DateSystem::V1900.base().date().and_hms(hour, min, sec)
    }
}

pub fn date_to_excel_number(date: impl ToDateTime, date_system: &DateSystem) -> f64 {
    let diff = date.to_datetime() - date_system.base();
    let adj = if *date_system == DateSystem::V1900 && diff.num_days() >= 60 { 1.0 } else { 0.0 };
    diff.num_milliseconds() as f64 / 1000.0 / 60.0 / 60.0 / 24.0 + adj
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

    #[test]
    fn v1900_num_to_date() {
        let expect = NaiveDate::from_ymd(1899, 12, 31).and_hms(0, 0, 0);
        match excel_number_to_date(0.0, &DateSystem::V1900) {
            Ok(date) => assert_eq!(date, expect),
            x => assert!(false, "did not convert 0.0 to proper date {:?}", x),
        }
    }

    #[test]
    fn v1900_num_after_bad_leap_to_date() {
        let expect = NaiveDate::from_ymd(1900, 3, 15).and_hms(0, 0, 0);
        match excel_number_to_date(75.0, &DateSystem::V1900) {
            Ok(date) => assert_eq!(date, expect),
            x => assert!(false, "did not convert 0.0 to proper date {:?}", x),
        }
    }

    #[test]
    fn v1900_num_with_time_date() {
        let expect = NaiveDate::from_ymd(1903, 5, 31).and_hms_milli(2, 17, 3, 34);
        match excel_number_to_date(1247.095174, &DateSystem::V1900) {
            Ok(date) => assert_eq!(date, expect),
            x => assert!(false, "did not convert 0.0 to proper date {:?}", x),
        }
    }

    #[test]
    fn v1900_date_to_num() {
        assert_eq!(0.0, date_to_excel_number(&NaiveDate::from_ymd(1899, 12, 31), &DateSystem::V1900));
    }

    #[test]
    fn v1900_bad_leap() {
        assert_eq!(61.0, date_to_excel_number(&NaiveDate::from_ymd(1900, 3, 1), &DateSystem::V1900));
    }

    #[test]
    fn v1900_before_bad_leap() {
        assert_eq!(59.0, date_to_excel_number(&NaiveDate::from_ymd(1900, 2, 28), &DateSystem::V1900));
    }

    #[test]
    fn v1900_with_time() {
        assert_eq!(128.5625, date_to_excel_number(&NaiveDate::from_ymd(1900, 5, 7).and_hms(13, 30, 0), &DateSystem::V1900));
    }

    #[test]
    fn v1904_date_to_num() {
        assert_eq!(0.0, date_to_excel_number(&NaiveDate::from_ymd(1904, 1, 1), &DateSystem::V1904));
    }
}
