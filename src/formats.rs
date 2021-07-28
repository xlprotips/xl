
mod parser {
    use std::str::Chars;

    // There will always be four formats, even though the user may only define one (or two or
    // three). See https://tinyurl.com/wrkptz2a for a thorough walkthrough, but succinctly:
    // - if 1 format provided, cover positive, negative, zero, and text
    // - if 2 provided, 1st = pos/zero/text, 2nd = neg
    // - if 3 provided, 1st = pos/text, 2nd = neg, 3rd = zero
    // - if 4 provided, 1st = pos, 2nd = neg, 3rd = zero, 4th = text
    struct Formats<'a> {
        positive: &'a str,
        negative: &'a str,
        zero: &'a str,
        text: &'a str,
    }

    #[derive(Debug)]
    enum Token {
        Asterisk,
        Backslash, // eg, \@ puts an @ before whatever comes next
        Char, // eg, \@ puts an @ before/after certain formats
        Colon, // Time
        Comma, // separate thousands, etc.
        Ident, // eg, General means no special formatting, mm = month format, etc.
        Period,
        PoundSign, // # has a special meaning
        QuestionMark,
        Str, // you can add strings before or after formats
        Underscore,
        Zero,
    }

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
            let mut chars = formula.chars();
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

        fn is_at_end(&self) -> bool {
            self.current.is_none()
        }

        fn advance(&mut self) -> Option<char> {
            self.current = self.peek;
            self.peek = self.chars.next();
            if let Some(c) = self.current {
                self.lexeme.push(c);
            }
            self.current
        }

        fn try_match(&mut self, expected: char) -> bool {
            if self.is_at_end() {
                return false
            }
            if self.peek != Some(expected) {
                return false
            }
            self.advance();
            true
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

        fn strip_lexeme(&mut self, c: char) -> String {
            self.lexeme.strip_prefix(c).unwrap_or(&self.lexeme).strip_suffix(c).unwrap_or(&self.lexeme).to_owned()
        }

        fn string(&mut self) -> Token {
            while let Some(c) = self.advance() {
                if c == '"' {
                    self.lexeme = self.strip_lexeme('"');
                    return self.token(TokenType::Str)
                }
            }
            self.error_msg("Unterminated string.".to_owned());
            self.token(TokenType::Unknown)
        }

        fn color(&mut self) -> Token {
            while let Some(c) = self.advance() {
                if c == ']' {
                    return self.token(Token::Color)
                }
            }
            self.error_msg("Unterminated range.".to_owned());
            self.token(TokenType::Unknown)
        }

        fn number(&mut self) -> Token {
            loop {
                let peek = self.peek();
                if peek == '#' || peek == ',' {
                    self.advance();
                } else {
                    break
                }
            }
            self.token(TokenType::Number)
        }

        fn ident(&mut self) -> Token {
            while self.peek().is_alphanumeric() { self.advance(); }
            self.token(TokenType::Ident)
        }
    }

    impl<'a> Iterator for Lexer<'a> {
        type Item = Token;
        fn next(&mut self) -> Option<Self::Item> {
            if let Some(c) = self.advance() {
                match c {
                    ',' => Some(self.token(TokenType::Comma)),
                    '.' => Some(self.token(TokenType::Period)),
                    '-' => Some(self.token(TokenType::Minus)),
                    '+' => Some(self.token(TokenType::Plus)),
                    ';' => Some(self.token(TokenType::Semicolon)),
                    '*' => Some(self.token(TokenType::Star)),
                    ':' => Some(self.token(TokenType::Colon)),
                    '#' => Some(self.token(TokenType::PoundSign)),
                    '0' => Some(self.token(TokenType::PoundSign)),
                    '"' => Some(self.string()),
                    '\'' => Some(self.path()),
                    '[' => Some(self.color()),
                    _ => {
                        self.error_msg(format!("Unexpected character: {}.", c));
                        Some(self.token(TokenType::Unknown))
                    }
                }
            } else {
                None
            }
        }
    }

}
