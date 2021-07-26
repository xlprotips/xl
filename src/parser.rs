use std::str::Chars;

#[derive(Debug)]
enum TokenType {
    Bang,
    BangEqual,
    Comma,
    Divide,
    Dot,
    Equal,
    EqualEqual,
    Greater,
    GreaterEqual,
    Ignore,
    LeftBrace,
    LeftParen,
    Less,
    LessEqual,
    Minus,
    Plus,
    RightBrace,
    RightParen,
    Semicolon,
    Star,
    Unknown,
}

#[derive(Debug)]
pub struct Token {
    index: usize,
    token_type: TokenType,
    value: String,
}

#[derive(Debug)]
pub struct Lexer<'a> {
    formula: &'a str,
    chars: Chars<'a>,
    current: Option<char>,
    peek: Option<char>,
    index: usize,
    line: usize,
    lexeme: String,
    had_error: bool,
}

impl Lexer<'_> {
    pub fn new(formula: &str) -> Lexer {
        let mut chars = formula.chars();
        let peek = chars.next();
        Lexer {
            formula,
            chars,
            current: None,
            peek,
            index: 1,
            line: 1,
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
        if self.current != Some(expected) {
            return false
        }
        self.advance();
        true
    }

    fn error(&mut self, msg: String) {
        self.had_error = true;
        eprintln!("[{}] {}", self.line, msg);
    }

    fn token(&mut self, token_type: TokenType) -> Token {
        let index = self.index;
        let value = self.lexeme.clone();
        self.lexeme.truncate(0);
        Token { index, token_type, value, }
    }

    fn peek(&self) -> char {
        self.peek.unwrap_or('\0')
    }
}

impl<'a> Iterator for Lexer<'a> {
    type Item = Token;
    fn next(&mut self) -> Option<Self::Item> {
        if let Some(c) = self.advance() {
            match c {
                '(' => Some(self.token(TokenType::LeftParen)),
                ')' => Some(self.token(TokenType::RightParen)),
                '{' => Some(self.token(TokenType::LeftBrace)),
                '}' => Some(self.token(TokenType::RightBrace)),
                ',' => Some(self.token(TokenType::Comma)),
                '.' => Some(self.token(TokenType::Dot)),
                '-' => Some(self.token(TokenType::Minus)),
                '+' => Some(self.token(TokenType::Plus)),
                ';' => Some(self.token(TokenType::Semicolon)),
                '*' => Some(self.token(TokenType::Star)),
                '/' => Some(self.token(TokenType::Divide)),
                '!' => {
                    if self.try_match('=') {
                        Some(self.token(TokenType::BangEqual))
                    } else {
                        Some(self.token(TokenType::Bang))
                    }
                },
                '=' => {
                    if self.try_match('=') {
                        Some(self.token(TokenType::EqualEqual))
                    } else {
                        Some(self.token(TokenType::Equal))
                    }
                },
                '<' => {
                    if self.try_match('=') {
                        Some(self.token(TokenType::LessEqual))
                    } else {
                        Some(self.token(TokenType::Less))
                    }
                },
                '>' => {
                    if self.try_match('=') {
                        Some(self.token(TokenType::GreaterEqual))
                    } else {
                        Some(self.token(TokenType::Greater))
                    }
                },
                ' ' => {
                    while self.peek() == ' ' {
                        self.advance();
                    }
                    Some(self.token(TokenType::Ignore))
                },
                _ => {
                    self.error(format!("Unexpected character: {}.", c));
                    Some(self.token(TokenType::Unknown))
                }
            }
        } else {
            None
        }
    }
}

/*

enum SubType {
    Start,
    Stop,
    Text,
    Number,
    Logical,
    Error,
    Range,
    Math,
    Concat,
    Intersect,
    Union,
}
*/

/*
fn get_tokens(formula: &str) {
    let mut formula = strip_formula(formula);
    let mut tokens: Vec<Token> = Vec::new();
    let mut tokenStack: Vec<Token> = Vec::new();
    let mut offset = 0;

    let eof = || offset >= formula.len();
    let next_char = || substring(formula, offset + 1, 1);
    let current_char = || substring(formula, offset, 1);
    let double_char = || substring(formula, offset, 2);

    let mut in_string = false;
    let mut in_path = false;
    let mut in_range = false;
    let mut in_error = false;

    while !eof() {
    }

}

/// Remove leading spaces and equal signs from formula
fn strip_formula(formula: &str) -> &str {
    let mut formula = formula;
    while formula.len() > 0 {
        let strip = |s| s == '=' || s == ' ';
        if let Some(stripped) = formula.strip_prefix(strip) {
            formula = stripped;
        }
    }
    formula
}
*/
