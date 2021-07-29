
mod parser {
    use std::str::Chars;

    // There will always be four formats, even though the user may only define one (or two or
    // three). See https://tinyurl.com/wrkptz2a for a thorough walkthrough, but succinctly:
    // - if 1 format provided, cover positive, negative, zero, and text
    // - if 2 provided, 1st = pos/zero/text, 2nd = neg
    // - if 3 provided, 1st = pos/text, 2nd = neg, 3rd = zero
    // - if 4 provided, 1st = pos, 2nd = neg, 3rd = zero, 4th = text
    /*
    struct Formats<'a> {
        positive: &'a str,
        negative: &'a str,
        zero: &'a str,
        text: &'a str,
    }
    */

    #[derive(Debug)]
    enum TokenType {
        // Special symbol - basically, no special format
        General,

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
    }

    impl Lexer<'_> {
        pub fn new(format: &str) -> Lexer {
            let mut chars = format.chars();
            let peek = chars.next();
            Lexer {
                format,
                chars,
                current: None,
                peek,
                index: 1,
                lexeme: String::new(),
                had_error: false,
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

    impl<'a> Iterator for Lexer<'a> {
        type Item = Token;
        fn next(&mut self) -> Option<Self::Item> {
            if let Some(c) = self.advance() {
                match c {
                    '0' => Some(self.token(TokenType::Zero)),
                    '#' => Some(self.token(TokenType::PoundSign)),
                    '?' => Some(self.token(TokenType::QuestionMark)),
                    ',' => Some(self.token(TokenType::Comma)),
                    '.' => Some(self.token(TokenType::Period)),
                    '/' => Some(self.token(TokenType::Slash)),
                    '%' => Some(self.token(TokenType::Percent)),
                    'e' | 'E' => Some(self.exponential()),
                    '*' => {
                        if self.peek() != '\0' {
                            self.advance();
                            self.lexeme = self.strip_lexeme('*');
                            Some(self.token(TokenType::Repeat))
                        } else {
                            dbg!("asterisk with no repeat");
                            Some(self.token(TokenType::Unknown))
                        }
                    },
                    '@' => Some(self.token(TokenType::At)),
                    '\'' => {
                        if self.peek() != '\0' {
                            self.advance();
                            self.lexeme = self.strip_lexeme('\'');
                            Some(self.token(TokenType::Text))
                        } else {
                            dbg!("asterisk with no repeat");
                            Some(self.token(TokenType::Unknown))
                        }
                    },
                    '_' => Some(self.token(TokenType::Underscore)),
                    '[' => {
                        match self.peek() {
                            '<' | '>' | '=' => Some(self.condition()),
                            _ => Some(self.color()),
                        }
                    },
                    'y' => Some(self.slurp_same(TokenType::Year)),
                    'm' => Some(self.slurp_same(TokenType::Month)),
                    'd' => Some(self.slurp_same(TokenType::Day)),
                    'h' => Some(self.slurp_same(TokenType::Hour)),
                    's' => Some(self.slurp_same(TokenType::Second)),
                    '"' => Some(self.string()),
                    'G' => {
                        for c in "eneral".chars() {
                            if !self.try_match(c) {
                                dbg!("expected 'General'");
                                return Some(self.token(TokenType::Unknown))
                            }
                        }
                        Some(self.token(TokenType::General))
                    },
                    'a' | 'A' => Some(self.time()),
                    ' ' => Some(self.slurp_same(TokenType::Text)),
                    _ => {
                        Some(self.token(TokenType::Text))
                    }
                }
            } else {
                None
            }
        }
    }

}

pub fn parse_format(format: &str) {
    let scanner = parser::Lexer::new(format);
    for token in scanner {
        println!("{:?}", token);
    }
}
