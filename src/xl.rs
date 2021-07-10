use std::convert::TryInto;
use regex::Regex;

const XL_MAX_COL: u16 = 16384;
const XL_MIN_COL: u16 = 1;


#[cfg(test)]
mod tests {
    mod utility_functions {
        use super::super::*;
        #[test]
        fn num_to_letter_w() {
            assert_eq!(col_num_to_letter(23), Some(String::from("W")));
        }

        #[test]
        fn num_to_letter_aa() {
            assert_eq!(col_num_to_letter(27), Some(String::from("AA")));
        }

        #[test]
        fn num_to_letter_ab() {
            assert_eq!(col_num_to_letter(28), Some(String::from("AB")));
        }

        #[test]
        fn num_to_letter_xfd() {
            assert_eq!(col_num_to_letter(16384), Some(String::from("XFD")));
        }

        #[test]
        fn num_to_letter_xfe() {
            assert_eq!(col_num_to_letter(16385), None);
        }

        #[test]
        fn num_to_letter_0() {
            assert_eq!(col_num_to_letter(0), None);
        }

        #[test]
        fn letter_to_num_w() {
            assert_eq!(col_letter_to_num("W"), Some(23));
        }

        #[test]
        fn letter_to_num_aa() {
            assert_eq!(col_letter_to_num("AA"), Some(27));
        }

        #[test]
        fn letter_to_num_ab() {
            assert_eq!(col_letter_to_num("AB"), Some(28));
        }

        #[test]
        fn letter_to_num_xfd() {
            assert_eq!(col_letter_to_num("XFD"), Some(16384));
        }

        #[test]
        fn letter_to_num_xfe() {
            assert_eq!(col_letter_to_num("XFE"), None);
        }

        #[test]
        fn letter_to_num_ab_lower() {
            assert_eq!(col_letter_to_num("ab"), Some(28));
        }

        #[test]
        fn letter_to_num_number() {
            assert_eq!(col_letter_to_num("12"), None);
        }

        #[test]
        fn letter_to_num_semicolon() {
            assert_eq!(col_letter_to_num(";"), None);
        }
    }
}


/// Return column letter for column number `n`
pub fn col_num_to_letter(n: u16) -> Option<String> {
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
pub fn col_letter_to_num(letter: &str) -> Option<u16> {
    let letter = letter.to_uppercase();
    let re = Regex::new(r"[A-Z]+").unwrap();
    if !re.is_match(&letter) { return None }
    let mut num: u16 = 0;
    for c in letter.chars() {
        num = num * 26 + ((c as u16) - ('A' as u16)) + 1;
    }
    if num > XL_MAX_COL || num < XL_MIN_COL { return None }
    Some(num)
}

