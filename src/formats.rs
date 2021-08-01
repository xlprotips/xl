use std::borrow::Cow;
use chrono::{NaiveDate, NaiveDateTime, NaiveTime, Timelike};

use crate::wb::DateSystem;

mod parser {
    use std::str::Chars;

    #[derive(Debug)]
    pub enum TokenType {
        // Special symbol - basically, no special format
        General,

        // there are (up to) four sections for number formats, broken up by semicolons.
        SectionBreak,

        // Number related codes
        Zero, // digit or zero
        PoundSign, // digit (if needed)
        Comma, // thousands separator
        Period, // show decimal point for numbers
        Slash, // fractions
        Percent, // multiply number by and add percent sign
        Exponential,
        QuestionMark, // digit or space
        Underscore, // skip width (like ?)

        // Text(ish) related codes
        Repeat, // *= means the "=" will fill empty space in the cell
        At, // @ ... when value is text, emit it "as is"
        Text, // Literal text to be emitted with value

        // Date-specific codes
        Year,
        Month,
        Day,
        Second,
        Hour,
        Meridiem, // AM/PM

        // "Special" codes
        Color, // eg, [Red]
        Condition, // eg, [<=100][Red] will display numbers less than 100 in red

        // Do not really expect to see this, but maybe?
        Unknown,
    }

    /*
    enum FormatColor {
        Black,
        Green,
        White,
        Blue,
        Magenta,
        Yellow,
        Cyan,
        Red,
    }
    */

    #[derive(Debug)]
    pub struct Token {
        index: usize,
        token_type: TokenType,
        value: String,
    }

    impl Token {
        pub fn token_type(&self) -> &TokenType {
            &self.token_type
        }
    }

    #[derive(Debug)]
    pub struct Lexer<'a> {
        // Total format
        format: &'a str,
        // format.chars() to help us iterate
        chars: Chars<'a>,
        // current char
        current: Option<char>,
        // next char
        peek: Option<char>,
        // current format code
        index: usize,
        // mutable string to help us keep track of format codes
        lexeme: String,
        // did we have any challenges parsing the format?
        had_error: bool,
        // list of tokens that we've seen so far
        tokens: Option<Vec<Token>>,
    }

    impl Lexer<'_> {
        pub fn new(format: &str) -> Lexer {
            let mut chars = format.chars();
            let peek = chars.next();
            let mut lexer = Lexer {
                format,
                chars,
                current: None,
                peek,
                index: 1,
                lexeme: String::new(),
                had_error: false,
                tokens: None,
            };
            lexer.prime();
            lexer
        }

        fn prime(&mut self) {
            let mut tokens = Vec::new();
            'main: loop {
                if let Some(c) = self.advance() {
                    let next_token = match c {
                        '0' => self.token(TokenType::Zero),
                        '#' => self.token(TokenType::PoundSign),
                        '?' => self.token(TokenType::QuestionMark),
                        ',' => {
                            self.token(TokenType::Comma)
                        },
                        '.' => self.token(TokenType::Period),
                        '/' => self.token(TokenType::Slash),
                        '%' => self.token(TokenType::Percent),
                        'e' | 'E' => self.exponential(),
                        '*' => {
                            if self.peek() != '\0' {
                                self.advance();
                                self.lexeme = self.strip_lexeme('*');
                                self.token(TokenType::Repeat)
                            } else {
                                dbg!("asterisk with no repeat");
                                self.token(TokenType::Unknown)
                            }
                        },
                        '@' => self.token(TokenType::At),
                        '\'' => {
                            if self.peek() != '\0' {
                                self.advance();
                                self.lexeme = self.strip_lexeme('\'');
                                self.token(TokenType::Text)
                            } else {
                                dbg!("asterisk with no repeat");
                                self.token(TokenType::Unknown)
                            }
                        },
                        '_' => self.token(TokenType::Underscore),
                        '[' => {
                            match self.peek() {
                                '<' | '>' | '=' => self.condition(),
                                _ => self.color(),
                            }
                        },
                        'y' => self.slurp_same(TokenType::Year),
                        'm' => self.slurp_same(TokenType::Month),
                        'd' => self.slurp_same(TokenType::Day),
                        'h' => self.slurp_same(TokenType::Hour),
                        's' => self.slurp_same(TokenType::Second),
                        '"' => self.string(),
                        'G' => {
                            for c in "eneral".chars() {
                                if !self.try_match(c) {
                                    dbg!("expected 'General'");
                                    tokens.push(self.token(TokenType::Unknown));
                                    continue 'main
                                }
                            }
                            self.token(TokenType::General)
                        },
                        'a' | 'A' => self.time(),
                        ' ' => self.slurp_same(TokenType::Text),
                        ';' => self.token(TokenType::SectionBreak),
                        _ => {
                            self.token(TokenType::Text)
                        }
                    };
                    tokens.push(next_token);
                } else {
                    self.tokens = Some(tokens);
                    return
                }
            }
        }

        fn advance(&mut self) -> Option<char> {
            self.current = self.peek;
            self.peek = self.chars.next();
            if let Some(c) = self.current {
                self.lexeme.push(c);
            }
            self.current
        }

        fn error_msg(&mut self, msg: String) {
            self.had_error = true;
            eprintln!("Error: {}", msg);
        }

        fn token(&mut self, token_type: TokenType) -> Token {
            let index = self.index;
            let value = self.lexeme.clone();
            self.lexeme.truncate(0);
            self.index += 1;
            Token { index, token_type, value, }
        }

        fn peek(&self) -> char {
            self.peek.unwrap_or('\0')
        }

        fn is_at_end(&self) -> bool {
            self.current.is_none()
        }

        fn try_match(&mut self, expected: char) -> bool {
            if self.is_at_end() { return false }
            if self.peek() != expected { return false }
            self.advance();
            true
        }

        fn strip_lexeme(&mut self, c: char) -> String {
            self.lexeme.strip_prefix(c).unwrap_or(&self.lexeme).strip_suffix(c).unwrap_or(&self.lexeme).to_owned()
        }

        fn string(&mut self) -> Token {
            while let Some(c) = self.advance() {
                if c == '"' {
                    self.lexeme = self.strip_lexeme('"');
                    return self.token(TokenType::Text)
                }
            }
            self.error_msg("Unterminated string.".to_owned());
            self.token(TokenType::Unknown)
        }

        fn color(&mut self) -> Token {
            while let Some(c) = self.advance() {
                if c == ']' {
                    let l = self.lexeme.strip_prefix('[').unwrap();
                    let l = l.strip_suffix(']').unwrap();
                    self.lexeme = l.to_owned();
                    return self.token(TokenType::Color)
                }
            }
            self.error_msg("Unterminated color.".to_owned());
            self.token(TokenType::Unknown)
        }

        fn condition(&mut self) -> Token {
            while let Some(c) = self.advance() {
                if c == ']' {
                    let lexeme = self.lexeme.strip_prefix('[').unwrap();
                    let lexeme = lexeme.strip_suffix(']').unwrap();
                    self.lexeme = lexeme.to_owned();
                    return self.token(TokenType::Condition)
                }
            }
            self.error_msg("Unterminated condition.".to_owned());
            self.token(TokenType::Unknown)
        }

        fn exponential(&mut self) -> Token {
            match self.peek() {
                '+' | '-' => { self.advance(); },
                _ => (),
            }
            self.token(TokenType::Exponential)
        }

        fn slurp_same(&mut self, token: TokenType) -> Token {
            while self.peek() == self.current.unwrap() {
                self.advance();
            }
            self.token(token)
        }

        fn time(&mut self) -> Token {
            match self.peek() {
                '/' => {
                    self.advance();
                    let peek = self.peek();
                    if peek == 'p' || peek == 'P' {
                        self.advance();
                        self.token(TokenType::Meridiem)
                    } else {
                        dbg!("expected p or P ending meridiem.");
                        self.token(TokenType::Unknown)
                    }
                },
                'm' | 'M' => {
                    self.advance();
                    if !self.try_match('/') {
                        dbg!("expected / to continue am/pm");
                        return self.token(TokenType::Unknown)
                    }
                    if self.peek() == 'P' || self.peek() == 'p' {
                        self.advance();
                    } else {
                        dbg!("expected 'p' to continue am/pm");
                        return self.token(TokenType::Unknown)
                    }
                    if self.peek() == 'm' || self.peek() == 'M' {
                        self.advance();
                    } else {
                        dbg!("expected 'm' to finish am/pm");
                        return self.token(TokenType::Unknown)
                    }
                    self.token(TokenType::Meridiem)
                },
                _ => {
                    dbg!("expected either '/' or 'm' to continue time");
                    self.token(TokenType::Unknown)
                }
            }
        }
    }

    impl IntoIterator for Lexer<'_> {
        type Item = Token;
        type IntoIter = ::std::vec::IntoIter<Token>;
        fn into_iter(self) -> Self::IntoIter {
            if let Some(tokens) = self.tokens {
                tokens.into_iter()
            } else {
                panic!("This shouldn't be possible");
            }
        }
    }

    impl<'a> IntoIterator for &'a Lexer<'_> {
        type Item = &'a Token;
        type IntoIter = ::std::slice::Iter<'a, Token>;
        fn into_iter(self) -> Self::IntoIter {
            if let Some(tokens) = &self.tokens {
                tokens.iter()
            } else {
                panic!("This shouldn't be possible");
            }
        }
    }

}

use crate::ExcelValue;
use parser::TokenType;

pub fn view_tokens(format: &str) {
    let scanner = parser::Lexer::new(format);
    for token in &scanner {
        println!("{:?}", token);
    }
}

struct Pad {
    with: char,
    n_times: usize,
}

impl Iterator for Pad {
    type Item = char;
    fn next(&mut self) -> Option<Self::Item> {
        if self.n_times > 0 {
            self.n_times -= 1;
            Some(self.with)
        } else {
            None
        }
    }
}

struct Formatter {
    number_of_required_digits: Option<usize>,
    extra_chars: Vec<(usize, String)>,
    show_commas: bool,
    number_of_decimals: Option<usize>,
}

fn format_number(num: &str, formatter: Formatter) -> String {
    println!("Formatting {}", num);
    let extra_chars = formatter.extra_chars;
    let mut extra_chars_idx = extra_chars.len();
    let mut formatted = String::new();
    let (whole, decimal) = if let Some(pos) = num.find('.') {
        num.split_at(pos)
    } else {
        (num, "")
    };
    let min_digits = formatter.number_of_required_digits.unwrap_or(0);
    let pad = Pad { with: '0', n_times: (min_digits - whole.len()).max(0) };
    let mut iorig = whole.len() + pad.n_times;
    for (i, c) in whole.chars().rev().chain(pad).enumerate() {
        if extra_chars_idx > 0 {
            let (pos, extra_char) = &extra_chars[extra_chars_idx-1];
            if *pos == iorig {
                extra_chars_idx -= 1;
                formatted.push_str(&extra_char);
            }
        }
        if formatter.show_commas && i != 0 && i % 3 == 0 {
            formatted.push(',');
        }
        formatted.push(c);
        iorig -= 1;
    }
    if extra_chars_idx > 0 {
        formatted.push_str(&extra_chars[extra_chars_idx-1].1);
    }
    formatted = formatted.chars().rev().collect();
    if let Some(n) = formatter.number_of_decimals {
        for c in decimal.chars().take(n) {
            formatted.push(c);
        }
    } else {
        for c in decimal.chars() {
            formatted.push(c);
        }
    }
    formatted
}

pub fn test_format_number(num: &str) {
    println!("Formatting {}", num);
    let extra_chars = vec![(0, "$"), (1, "k"), (9, "?")];
    let mut extra_chars_idx = extra_chars.len();
    let mut formatted = String::new();
    let (whole, decimal) = if let Some(pos) = num.find('.') {
        num.split_at(pos)
    } else {
        (num, "")
    };
    let min_digits = 9;
    let pad = Pad { with: '0', n_times: (min_digits - whole.len()).max(0) };
    let mut iorig = whole.len() + pad.n_times;
    for (i, c) in whole.chars().rev().chain(pad).enumerate() {
        if extra_chars_idx > 0 {
            let (pos, extra_char) = extra_chars[extra_chars_idx-1];
            if pos == iorig {
                extra_chars_idx -= 1;
                formatted.push_str(extra_char);
            }
        }
        if i != 0 && i % 3 == 0 {
            formatted.push(',');
        }
        formatted.push(c);
        iorig -= 1;
    }
    if extra_chars_idx > 0 {
        formatted.push_str(extra_chars[extra_chars_idx-1].1);
    }
    formatted = formatted.chars().rev().collect();
    for c in decimal.chars() {
        formatted.push(c);
    }
    println!("{}", formatted);
}

// There will always be four formats, even though the user may only define one (or two or
// three). See https://tinyurl.com/wrkptz2a for a thorough walkthrough, but succinctly:
// - if 1 format provided, cover positive, negative, zero, and text
// - if 2 provided, 1st = pos/zero/text, 2nd = neg
// - if 3 provided, 1st = pos/text, 2nd = neg, 3rd = zero
// - if 4 provided, 1st = pos, 2nd = neg, 3rd = zero, 4th = text
fn parse_format(format: &str) -> impl FnOnce(&ExcelValue) -> String {
    let scanner = parser::Lexer::new(format);
    let mut number_formatter = Formatter {
        number_of_required_digits: None,
        extra_chars: vec![],
        show_commas: false,
        number_of_decimals: None,
    };
    let mut seen_period = false;
    for token in scanner {
        match token.token_type() {
            TokenType::Zero => {
                if seen_period {
                    let n = number_formatter.number_of_required_digits.get_or_insert(0);
                    *n += 1;
                } else {
                    let n = number_formatter.number_of_decimals.get_or_insert(0);
                    *n += 1;
                }
            },
            TokenType::PoundSign => (),
            TokenType::Comma => {
                if number_formatter.number_of_required_digits.is_some() {
                    number_formatter.show_commas = true;
                } else {
                    // number_formatter.extra_chars[]
                }
            },
            TokenType::Period => seen_period = true,
            TokenType::Slash => (),
            TokenType::Percent => (),
            TokenType::Exponential => (),
            TokenType::QuestionMark => (),
            TokenType::Underscore => (),
            _ => (),
        }
    }
    let formatter = move |v: &ExcelValue| {
        let string = String::from(v);
        format_number(&string, number_formatter)
    };
    formatter
}

impl ExcelValue<'_> {
    pub fn format(&self, with: &str) -> String {
        let formatter = parse_format(with);
        formatter(&self)
    }
}

pub trait ToExcelValue {
    fn to_excel(&self) -> ExcelValue;
}

pub fn format(value: impl ToExcelValue, with: &str) -> String {
    let v = value.to_excel();
    v.format(with)
}

impl ToExcelValue for bool {
    fn to_excel(&self) -> ExcelValue { ExcelValue::Bool(*self) }
}

impl ToExcelValue for &str {
    fn to_excel(&self) -> ExcelValue { ExcelValue::String(Cow::Borrowed(self)) }
}

impl ToExcelValue for String {
    fn to_excel(&self) -> ExcelValue { ExcelValue::String(Cow::Borrowed(&self)) }
}

impl ToExcelValue for NaiveDate {
    fn to_excel(&self) -> ExcelValue {
        let num = crate::date_to_excel_number(self, &DateSystem::V1900);
        ExcelValue::Date(self.and_hms(0, 0, 0), num)
    }
}

impl ToExcelValue for NaiveDateTime {
    fn to_excel(&self) -> ExcelValue {
        let num = crate::date_to_excel_number(self, &DateSystem::V1900);
        ExcelValue::DateTime(*self, num)
    }
}

impl ToExcelValue for NaiveTime {
    fn to_excel(&self) -> ExcelValue {
        let num = crate::date_to_excel_number(self, &DateSystem::V1900);
        let date = NaiveDate::from_ymd(1899, 12, 31).and_hms(self.hour(), self.minute(), self.second());
        ExcelValue::Time(date, num)
    }
}

impl ToExcelValue for f64 {
    fn to_excel(&self) -> ExcelValue { ExcelValue::Number(*self) }
}

impl ToExcelValue for i32 {
    fn to_excel(&self) -> ExcelValue { ExcelValue::Number(f64::from(*self)) }
}
