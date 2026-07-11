//! A spreadsheet formula evaluator covering the everyday function set, so
//! agents can read computed results back without opening Excel.
//!
//! Scope: arithmetic/comparison/concatenation operators, cell and range
//! references (cross-sheet, absolute/relative), and ~40 common functions.
//! Formulas referencing other formula cells evaluate recursively with
//! cycle detection. Anything unsupported returns a #NAME?/#VALUE!-style
//! error value rather than failing the command.

use std::cell::RefCell;
use std::collections::HashSet;

use crate::path::parse_cell_ref;

/// Raw cell content as stored in the sheet.
#[derive(Debug, Clone)]
pub enum CellContent {
    Empty,
    Number(f64),
    Text(String),
    Bool(bool),
    Formula(String),
    Error(String),
}

/// Where the evaluator reads cells from.
pub trait CellSource {
    /// `sheet` is always fully qualified by the evaluator.
    fn cell(&self, sheet: &str, col: u32, row: u32) -> CellContent;
    fn has_sheet(&self, sheet: &str) -> bool;
}

/// A computed value.
#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Number(f64),
    Text(String),
    Bool(bool),
    Empty,
    Err(String),
}

impl Value {
    pub fn display(&self) -> String {
        match self {
            Value::Number(n) => {
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    format!("{}", *n as i64)
                } else {
                    let s = format!("{:.10}", n);
                    let s = s.trim_end_matches('0').trim_end_matches('.');
                    s.to_string()
                }
            }
            Value::Text(t) => t.clone(),
            Value::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
            Value::Empty => String::new(),
            Value::Err(e) => e.clone(),
        }
    }

    pub fn type_name(&self) -> &'static str {
        match self {
            Value::Number(_) => "number",
            Value::Text(_) => "string",
            Value::Bool(_) => "boolean",
            Value::Empty => "empty",
            Value::Err(_) => "error",
        }
    }

    fn as_number(&self) -> Result<f64, Value> {
        match self {
            Value::Number(n) => Ok(*n),
            Value::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
            Value::Empty => Ok(0.0),
            Value::Text(t) => t
                .trim()
                .parse::<f64>()
                .map_err(|_| Value::Err("#VALUE!".into())),
            Value::Err(_) => Err(self.clone()),
        }
    }

    fn as_text(&self) -> String {
        match self {
            Value::Empty => String::new(),
            other => other.display(),
        }
    }

    fn as_bool(&self) -> Result<bool, Value> {
        match self {
            Value::Bool(b) => Ok(*b),
            Value::Number(n) => Ok(*n != 0.0),
            Value::Empty => Ok(false),
            Value::Text(t) if t.eq_ignore_ascii_case("true") => Ok(true),
            Value::Text(t) if t.eq_ignore_ascii_case("false") => Ok(false),
            Value::Text(_) => Err(Value::Err("#VALUE!".into())),
            Value::Err(_) => Err(self.clone()),
        }
    }
}

// ---------------------------------------------------------------- AST ----

#[derive(Debug, Clone, PartialEq)]
enum Op {
    Add,
    Sub,
    Mul,
    Div,
    Pow,
    Concat,
    Eq,
    Ne,
    Lt,
    Gt,
    Le,
    Ge,
}

#[derive(Debug, Clone)]
enum Expr {
    Num(f64),
    Str(String),
    Bool(bool),
    Ref {
        sheet: Option<String>,
        col: u32,
        row: u32,
    },
    Range {
        sheet: Option<String>,
        c0: u32,
        r0: u32,
        c1: u32,
        r1: u32,
    },
    Func(String, Vec<Expr>),
    Bin(Op, Box<Expr>, Box<Expr>),
    Neg(Box<Expr>),
    Percent(Box<Expr>),
}

// ---------------------------------------------------------- tokenizer ----

#[derive(Debug, Clone, PartialEq)]
enum Tok {
    Num(f64),
    Str(String),
    /// Identifier / cell ref fragment (letters, digits, $, _, .).
    Ident(String),
    /// 'Quoted Sheet Name'
    QuotedIdent(String),
    Op(Op),
    LParen,
    RParen,
    Comma,
    Colon,
    Bang,
    Percent,
}

fn tokenize(input: &str) -> Result<Vec<Tok>, String> {
    let mut out = Vec::new();
    let mut chars = input.chars().peekable();
    while let Some(&c) = chars.peek() {
        match c {
            ' ' | '\t' | '\n' | '\r' => {
                chars.next();
            }
            '0'..='9' | '.' => {
                let mut s = String::new();
                while let Some(&d) = chars.peek() {
                    if d.is_ascii_digit() || d == '.' {
                        s.push(d);
                        chars.next();
                    } else if (d == 'e' || d == 'E')
                        && s.chars().any(|c| c.is_ascii_digit())
                        && matches!(
                            chars.clone().nth(1),
                            Some(c2) if c2.is_ascii_digit() || c2 == '-' || c2 == '+'
                        )
                    {
                        s.push(d);
                        chars.next();
                        if let Some(&sign) = chars.peek() {
                            if sign == '-' || sign == '+' {
                                s.push(sign);
                                chars.next();
                            }
                        }
                    } else {
                        break;
                    }
                }
                out.push(Tok::Num(s.parse().map_err(|_| format!("bad number '{s}'"))?));
            }
            '"' => {
                chars.next();
                let mut s = String::new();
                loop {
                    match chars.next() {
                        Some('"') => {
                            if chars.peek() == Some(&'"') {
                                s.push('"');
                                chars.next();
                            } else {
                                break;
                            }
                        }
                        Some(ch) => s.push(ch),
                        None => return Err("unterminated string".into()),
                    }
                }
                out.push(Tok::Str(s));
            }
            '\'' => {
                chars.next();
                let mut s = String::new();
                loop {
                    match chars.next() {
                        Some('\'') => {
                            if chars.peek() == Some(&'\'') {
                                s.push('\'');
                                chars.next();
                            } else {
                                break;
                            }
                        }
                        Some(ch) => s.push(ch),
                        None => return Err("unterminated sheet name".into()),
                    }
                }
                out.push(Tok::QuotedIdent(s));
            }
            'A'..='Z' | 'a'..='z' | '$' | '_' => {
                let mut s = String::new();
                while let Some(&d) = chars.peek() {
                    if d.is_ascii_alphanumeric() || d == '$' || d == '_' || d == '.' {
                        s.push(d);
                        chars.next();
                    } else {
                        break;
                    }
                }
                out.push(Tok::Ident(s));
            }
            '(' => {
                chars.next();
                out.push(Tok::LParen);
            }
            ')' => {
                chars.next();
                out.push(Tok::RParen);
            }
            ',' | ';' => {
                chars.next();
                out.push(Tok::Comma);
            }
            ':' => {
                chars.next();
                out.push(Tok::Colon);
            }
            '!' => {
                chars.next();
                out.push(Tok::Bang);
            }
            '%' => {
                chars.next();
                out.push(Tok::Percent);
            }
            '+' => {
                chars.next();
                out.push(Tok::Op(Op::Add));
            }
            '-' => {
                chars.next();
                out.push(Tok::Op(Op::Sub));
            }
            '*' => {
                chars.next();
                out.push(Tok::Op(Op::Mul));
            }
            '/' => {
                chars.next();
                out.push(Tok::Op(Op::Div));
            }
            '^' => {
                chars.next();
                out.push(Tok::Op(Op::Pow));
            }
            '&' => {
                chars.next();
                out.push(Tok::Op(Op::Concat));
            }
            '=' => {
                chars.next();
                out.push(Tok::Op(Op::Eq));
            }
            '<' => {
                chars.next();
                match chars.peek() {
                    Some('=') => {
                        chars.next();
                        out.push(Tok::Op(Op::Le));
                    }
                    Some('>') => {
                        chars.next();
                        out.push(Tok::Op(Op::Ne));
                    }
                    _ => out.push(Tok::Op(Op::Lt)),
                }
            }
            '>' => {
                chars.next();
                if chars.peek() == Some(&'=') {
                    chars.next();
                    out.push(Tok::Op(Op::Ge));
                } else {
                    out.push(Tok::Op(Op::Gt));
                }
            }
            other => return Err(format!("unexpected character '{other}'")),
        }
    }
    Ok(out)
}

// ------------------------------------------------------------- parser ----

struct Parser {
    toks: Vec<Tok>,
    pos: usize,
}

impl Parser {
    fn peek(&self) -> Option<&Tok> {
        self.toks.get(self.pos)
    }

    fn next(&mut self) -> Option<Tok> {
        let t = self.toks.get(self.pos).cloned();
        if t.is_some() {
            self.pos += 1;
        }
        t
    }

    fn expect(&mut self, t: &Tok) -> Result<(), String> {
        match self.next() {
            Some(ref got) if got == t => Ok(()),
            got => Err(format!("expected {t:?}, got {got:?}")),
        }
    }

    /// comparison ( = <> < > <= >= )
    fn expr(&mut self) -> Result<Expr, String> {
        let mut left = self.concat()?;
        while let Some(Tok::Op(op)) = self.peek() {
            let op = match op {
                Op::Eq | Op::Ne | Op::Lt | Op::Gt | Op::Le | Op::Ge => op.clone(),
                _ => break,
            };
            self.next();
            let right = self.concat()?;
            left = Expr::Bin(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn concat(&mut self) -> Result<Expr, String> {
        let mut left = self.additive()?;
        while matches!(self.peek(), Some(Tok::Op(Op::Concat))) {
            self.next();
            let right = self.additive()?;
            left = Expr::Bin(Op::Concat, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn additive(&mut self) -> Result<Expr, String> {
        let mut left = self.multiplicative()?;
        while let Some(Tok::Op(op @ (Op::Add | Op::Sub))) = self.peek() {
            let op = op.clone();
            self.next();
            let right = self.multiplicative()?;
            left = Expr::Bin(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn multiplicative(&mut self) -> Result<Expr, String> {
        let mut left = self.power()?;
        while let Some(Tok::Op(op @ (Op::Mul | Op::Div))) = self.peek() {
            let op = op.clone();
            self.next();
            let right = self.power()?;
            left = Expr::Bin(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn power(&mut self) -> Result<Expr, String> {
        let base = self.unary()?;
        if matches!(self.peek(), Some(Tok::Op(Op::Pow))) {
            self.next();
            // Right-associative.
            let exp = self.power()?;
            return Ok(Expr::Bin(Op::Pow, Box::new(base), Box::new(exp)));
        }
        Ok(base)
    }

    fn unary(&mut self) -> Result<Expr, String> {
        match self.peek() {
            Some(Tok::Op(Op::Sub)) => {
                self.next();
                Ok(Expr::Neg(Box::new(self.unary()?)))
            }
            Some(Tok::Op(Op::Add)) => {
                self.next();
                self.unary()
            }
            _ => self.postfix(),
        }
    }

    fn postfix(&mut self) -> Result<Expr, String> {
        let mut e = self.primary()?;
        while matches!(self.peek(), Some(Tok::Percent)) {
            self.next();
            e = Expr::Percent(Box::new(e));
        }
        Ok(e)
    }

    fn primary(&mut self) -> Result<Expr, String> {
        match self.next() {
            Some(Tok::Num(n)) => Ok(Expr::Num(n)),
            Some(Tok::Str(s)) => Ok(Expr::Str(s)),
            Some(Tok::LParen) => {
                let e = self.expr()?;
                self.expect(&Tok::RParen)?;
                Ok(e)
            }
            Some(Tok::QuotedIdent(sheet)) => {
                self.expect(&Tok::Bang)?;
                self.reference(Some(sheet))
            }
            Some(Tok::Ident(word)) => {
                // Sheet!ref, function call, boolean, or a bare cell ref.
                if matches!(self.peek(), Some(Tok::Bang)) {
                    self.next();
                    return self.reference(Some(word));
                }
                if matches!(self.peek(), Some(Tok::LParen)) {
                    self.next();
                    let name = word.to_ascii_uppercase();
                    let mut args = Vec::new();
                    if !matches!(self.peek(), Some(Tok::RParen)) {
                        loop {
                            args.push(self.expr()?);
                            match self.next() {
                                Some(Tok::Comma) => continue,
                                Some(Tok::RParen) => break,
                                got => return Err(format!("expected , or ) in {name}(), got {got:?}")),
                            }
                        }
                    } else {
                        self.next();
                    }
                    return Ok(Expr::Func(name, args));
                }
                match word.to_ascii_uppercase().as_str() {
                    "TRUE" => Ok(Expr::Bool(true)),
                    "FALSE" => Ok(Expr::Bool(false)),
                    _ => self.reference_from(None, word),
                }
            }
            got => Err(format!("unexpected token {got:?}")),
        }
    }

    /// Parse the ref/range that follows `Sheet!`.
    fn reference(&mut self, sheet: Option<String>) -> Result<Expr, String> {
        match self.next() {
            Some(Tok::Ident(word)) => self.reference_from(sheet, word),
            got => Err(format!("expected a cell reference, got {got:?}")),
        }
    }

    fn reference_from(&mut self, sheet: Option<String>, word: String) -> Result<Expr, String> {
        let clean = word.replace('$', "");
        let Some((c0, r0)) = parse_cell_ref(&clean) else {
            return Err(format!("#NAME? '{word}'"));
        };
        if matches!(self.peek(), Some(Tok::Colon)) {
            self.next();
            let end = match self.next() {
                Some(Tok::Ident(w)) => w,
                got => return Err(format!("expected range end, got {got:?}")),
            };
            let clean_end = end.replace('$', "");
            let Some((c1, r1)) = parse_cell_ref(&clean_end) else {
                return Err(format!("#NAME? '{end}'"));
            };
            return Ok(Expr::Range {
                sheet,
                c0: c0.min(c1),
                r0: r0.min(r1),
                c1: c0.max(c1),
                r1: r0.max(r1),
            });
        }
        Ok(Expr::Ref { sheet, col: c0, row: r0 })
    }
}

fn parse(formula: &str) -> Result<Expr, String> {
    let body = formula.strip_prefix('=').unwrap_or(formula);
    let toks = tokenize(body)?;
    let mut p = Parser { toks, pos: 0 };
    let e = p.expr()?;
    if p.pos != p.toks.len() {
        return Err(format!("trailing tokens after formula: {:?}", &p.toks[p.pos..]));
    }
    Ok(e)
}

// ---------------------------------------------------------- evaluator ----

const MAX_DEPTH: usize = 64;

pub struct Evaluator<'a> {
    source: &'a dyn CellSource,
    visiting: RefCell<HashSet<(String, u32, u32)>>,
    depth: RefCell<usize>,
}

impl<'a> Evaluator<'a> {
    pub fn new(source: &'a dyn CellSource) -> Evaluator<'a> {
        Evaluator {
            source,
            visiting: RefCell::new(HashSet::new()),
            depth: RefCell::new(0),
        }
    }

    /// Evaluate a formula in the context of `sheet` (for unqualified refs).
    pub fn eval_formula(&self, sheet: &str, formula: &str) -> Value {
        {
            let mut d = self.depth.borrow_mut();
            if *d >= MAX_DEPTH {
                return Value::Err("#DEPTH!".into());
            }
            *d += 1;
        }
        let result = match parse(formula) {
            Ok(expr) => self.eval(sheet, &expr),
            Err(e) => {
                if e.starts_with("#NAME?") {
                    Value::Err("#NAME?".into())
                } else {
                    Value::Err(format!("#PARSE! ({e})"))
                }
            }
        };
        *self.depth.borrow_mut() -= 1;
        result
    }

    fn eval_cell(&self, sheet: &str, col: u32, row: u32) -> Value {
        let key = (sheet.to_lowercase(), col, row);
        if self.visiting.borrow().contains(&key) {
            return Value::Err("#CIRC!".into());
        }
        match self.source.cell(sheet, col, row) {
            CellContent::Empty => Value::Empty,
            CellContent::Number(n) => Value::Number(n),
            CellContent::Text(t) => Value::Text(t),
            CellContent::Bool(b) => Value::Bool(b),
            CellContent::Error(e) => Value::Err(e),
            CellContent::Formula(f) => {
                self.visiting.borrow_mut().insert(key.clone());
                let v = self.eval_formula(sheet, &f);
                self.visiting.borrow_mut().remove(&key);
                v
            }
        }
    }

    /// Flatten an argument into scalar values (ranges expand).
    fn flatten(&self, sheet: &str, e: &Expr, out: &mut Vec<Value>) -> Result<(), Value> {
        match e {
            Expr::Range { sheet: s, c0, r0, c1, r1 } => {
                let sname = s.as_deref().unwrap_or(sheet);
                if !self.source.has_sheet(sname) {
                    return Err(Value::Err("#REF!".into()));
                }
                if (r1 - r0) as u64 * (c1 - c0) as u64 > 1_000_000 {
                    return Err(Value::Err("#RANGE-TOO-BIG!".into()));
                }
                for r in *r0..=*r1 {
                    for c in *c0..=*c1 {
                        out.push(self.eval_cell(sname, c, r));
                    }
                }
                Ok(())
            }
            other => {
                out.push(self.eval(sheet, other));
                Ok(())
            }
        }
    }

    /// Range dimensions + values for lookup functions (row-major).
    fn range_grid(&self, sheet: &str, e: &Expr) -> Result<(usize, usize, Vec<Value>), Value> {
        match e {
            Expr::Range { sheet: s, c0, r0, c1, r1 } => {
                let sname = s.as_deref().unwrap_or(sheet);
                if !self.source.has_sheet(sname) {
                    return Err(Value::Err("#REF!".into()));
                }
                let (w, h) = ((c1 - c0 + 1) as usize, (r1 - r0 + 1) as usize);
                if w * h > 1_000_000 {
                    return Err(Value::Err("#RANGE-TOO-BIG!".into()));
                }
                let mut vals = Vec::with_capacity(w * h);
                for r in *r0..=*r1 {
                    for c in *c0..=*c1 {
                        vals.push(self.eval_cell(sname, c, r));
                    }
                }
                Ok((w, h, vals))
            }
            _ => Err(Value::Err("#VALUE! (expected a range)".into())),
        }
    }

    fn eval(&self, sheet: &str, e: &Expr) -> Value {
        match e {
            Expr::Num(n) => Value::Number(*n),
            Expr::Str(s) => Value::Text(s.clone()),
            Expr::Bool(b) => Value::Bool(*b),
            Expr::Ref { sheet: s, col, row } => {
                let sname = s.as_deref().unwrap_or(sheet);
                if !self.source.has_sheet(sname) {
                    return Value::Err("#REF!".into());
                }
                self.eval_cell(sname, *col, *row)
            }
            Expr::Range { .. } => Value::Err("#VALUE! (range used as a scalar)".into()),
            Expr::Neg(inner) => match self.eval(sheet, inner).as_number() {
                Ok(n) => Value::Number(-n),
                Err(e) => e,
            },
            Expr::Percent(inner) => match self.eval(sheet, inner).as_number() {
                Ok(n) => Value::Number(n / 100.0),
                Err(e) => e,
            },
            Expr::Bin(op, l, r) => self.eval_bin(sheet, op, l, r),
            Expr::Func(name, args) => self.eval_func(sheet, name, args),
        }
    }

    fn eval_bin(&self, sheet: &str, op: &Op, l: &Expr, r: &Expr) -> Value {
        let lv = self.eval(sheet, l);
        let rv = self.eval(sheet, r);
        if let Value::Err(_) = lv {
            return lv;
        }
        if let Value::Err(_) = rv {
            return rv;
        }
        match op {
            Op::Concat => Value::Text(format!("{}{}", lv.as_text(), rv.as_text())),
            Op::Add | Op::Sub | Op::Mul | Op::Div | Op::Pow => {
                let a = match lv.as_number() {
                    Ok(n) => n,
                    Err(e) => return e,
                };
                let b = match rv.as_number() {
                    Ok(n) => n,
                    Err(e) => return e,
                };
                match op {
                    Op::Add => Value::Number(a + b),
                    Op::Sub => Value::Number(a - b),
                    Op::Mul => Value::Number(a * b),
                    Op::Div => {
                        if b == 0.0 {
                            Value::Err("#DIV/0!".into())
                        } else {
                            Value::Number(a / b)
                        }
                    }
                    Op::Pow => Value::Number(a.powf(b)),
                    _ => unreachable!(),
                }
            }
            Op::Eq | Op::Ne | Op::Lt | Op::Gt | Op::Le | Op::Ge => {
                let ord = compare_values(&lv, &rv);
                let b = match op {
                    Op::Eq => ord == std::cmp::Ordering::Equal,
                    Op::Ne => ord != std::cmp::Ordering::Equal,
                    Op::Lt => ord == std::cmp::Ordering::Less,
                    Op::Gt => ord == std::cmp::Ordering::Greater,
                    Op::Le => ord != std::cmp::Ordering::Greater,
                    Op::Ge => ord != std::cmp::Ordering::Less,
                    _ => unreachable!(),
                };
                Value::Bool(b)
            }
        }
    }

    fn eval_func(&self, sheet: &str, name: &str, args: &[Expr]) -> Value {
        macro_rules! numbers {
            () => {{
                let mut vals = Vec::new();
                for a in args {
                    if let Err(e) = self.flatten(sheet, a, &mut vals) {
                        return e;
                    }
                }
                let mut nums = Vec::new();
                for v in vals {
                    match v {
                        Value::Number(n) => nums.push(n),
                        Value::Bool(_) | Value::Text(_) | Value::Empty => {} // aggregates skip non-numbers
                        Value::Err(_) => return v,
                    }
                }
                nums
            }};
        }
        macro_rules! arg_num {
            ($i:expr) => {
                match self.eval(sheet, &args[$i]).as_number() {
                    Ok(n) => n,
                    Err(e) => return e,
                }
            };
        }
        macro_rules! arg_text {
            ($i:expr) => {{
                let v = self.eval(sheet, &args[$i]);
                if let Value::Err(_) = v {
                    return v;
                }
                v.as_text()
            }};
        }
        macro_rules! need {
            ($n:expr) => {
                if args.len() < $n {
                    return Value::Err(format!("#N/A ({name} needs {} argument(s))", $n));
                }
            };
        }

        match name {
            "SUM" => Value::Number(numbers!().iter().sum()),
            "PRODUCT" => Value::Number(numbers!().iter().product()),
            "AVERAGE" => {
                let nums = numbers!();
                if nums.is_empty() {
                    Value::Err("#DIV/0!".into())
                } else {
                    Value::Number(nums.iter().sum::<f64>() / nums.len() as f64)
                }
            }
            "MIN" => Value::Number(numbers!().into_iter().fold(f64::INFINITY, f64::min).min(f64::INFINITY)).map_empty(),
            "MAX" => Value::Number(numbers!().into_iter().fold(f64::NEG_INFINITY, f64::max)).map_empty(),
            "COUNT" => Value::Number(numbers!().len() as f64),
            "COUNTA" => {
                let mut vals = Vec::new();
                for a in args {
                    if let Err(e) = self.flatten(sheet, a, &mut vals) {
                        return e;
                    }
                }
                Value::Number(vals.iter().filter(|v| !matches!(v, Value::Empty)).count() as f64)
            }
            "MEDIAN" => {
                let mut nums = numbers!();
                if nums.is_empty() {
                    return Value::Err("#NUM!".into());
                }
                nums.sort_by(|a, b| a.partial_cmp(b).unwrap());
                let mid = nums.len() / 2;
                Value::Number(if nums.len() % 2 == 0 {
                    (nums[mid - 1] + nums[mid]) / 2.0
                } else {
                    nums[mid]
                })
            }
            "IF" => {
                need!(2);
                match self.eval(sheet, &args[0]).as_bool() {
                    Ok(true) => self.eval(sheet, &args[1]),
                    Ok(false) => {
                        if args.len() > 2 {
                            self.eval(sheet, &args[2])
                        } else {
                            Value::Bool(false)
                        }
                    }
                    Err(e) => e,
                }
            }
            "IFERROR" => {
                need!(2);
                let v = self.eval(sheet, &args[0]);
                if matches!(v, Value::Err(_)) {
                    self.eval(sheet, &args[1])
                } else {
                    v
                }
            }
            "AND" | "OR" => {
                let mut vals = Vec::new();
                for a in args {
                    if let Err(e) = self.flatten(sheet, a, &mut vals) {
                        return e;
                    }
                }
                let mut acc = name == "AND";
                for v in vals {
                    match v.as_bool() {
                        Ok(b) => {
                            if name == "AND" {
                                acc &= b;
                            } else {
                                acc |= b;
                            }
                        }
                        Err(Value::Err(e)) => return Value::Err(e),
                        Err(_) => {}
                    }
                }
                Value::Bool(acc)
            }
            "NOT" => {
                need!(1);
                match self.eval(sheet, &args[0]).as_bool() {
                    Ok(b) => Value::Bool(!b),
                    Err(e) => e,
                }
            }
            "ROUND" | "ROUNDUP" | "ROUNDDOWN" => {
                need!(1);
                let n = arg_num!(0);
                let digits = if args.len() > 1 { arg_num!(1) } else { 0.0 };
                let factor = 10f64.powf(digits.trunc());
                let scaled = n * factor;
                let rounded = match name {
                    "ROUNDUP" => scaled.abs().ceil() * scaled.signum(),
                    "ROUNDDOWN" => scaled.trunc(),
                    _ => scaled.round(),
                };
                Value::Number(rounded / factor)
            }
            "INT" => {
                need!(1);
                Value::Number(arg_num!(0).floor())
            }
            "ABS" => {
                need!(1);
                Value::Number(arg_num!(0).abs())
            }
            "MOD" => {
                need!(2);
                let (a, b) = (arg_num!(0), arg_num!(1));
                if b == 0.0 {
                    Value::Err("#DIV/0!".into())
                } else {
                    Value::Number(a - b * (a / b).floor())
                }
            }
            "POWER" => {
                need!(2);
                Value::Number(arg_num!(0).powf(arg_num!(1)))
            }
            "SQRT" => {
                need!(1);
                let n = arg_num!(0);
                if n < 0.0 {
                    Value::Err("#NUM!".into())
                } else {
                    Value::Number(n.sqrt())
                }
            }
            "EXP" => {
                need!(1);
                Value::Number(arg_num!(0).exp())
            }
            "LN" => {
                need!(1);
                Value::Number(arg_num!(0).ln())
            }
            "LOG10" => {
                need!(1);
                Value::Number(arg_num!(0).log10())
            }
            "CONCAT" | "CONCATENATE" => {
                let mut vals = Vec::new();
                for a in args {
                    if let Err(e) = self.flatten(sheet, a, &mut vals) {
                        return e;
                    }
                }
                let mut s = String::new();
                for v in vals {
                    if let Value::Err(_) = v {
                        return v;
                    }
                    s.push_str(&v.as_text());
                }
                Value::Text(s)
            }
            "LEFT" | "RIGHT" => {
                need!(1);
                let t = arg_text!(0);
                let n = if args.len() > 1 { arg_num!(1).max(0.0) as usize } else { 1 };
                let chars: Vec<char> = t.chars().collect();
                let n = n.min(chars.len());
                let s: String = if name == "LEFT" {
                    chars[..n].iter().collect()
                } else {
                    chars[chars.len() - n..].iter().collect()
                };
                Value::Text(s)
            }
            "MID" => {
                need!(3);
                let t = arg_text!(0);
                let start = (arg_num!(1).max(1.0) as usize).saturating_sub(1);
                let len = arg_num!(2).max(0.0) as usize;
                let chars: Vec<char> = t.chars().collect();
                let start = start.min(chars.len());
                let end = (start + len).min(chars.len());
                Value::Text(chars[start..end].iter().collect())
            }
            "LEN" => {
                need!(1);
                Value::Number(arg_text!(0).chars().count() as f64)
            }
            "UPPER" => {
                need!(1);
                Value::Text(arg_text!(0).to_uppercase())
            }
            "LOWER" => {
                need!(1);
                Value::Text(arg_text!(0).to_lowercase())
            }
            "TRIM" => {
                need!(1);
                let t = arg_text!(0);
                Value::Text(t.split_whitespace().collect::<Vec<_>>().join(" "))
            }
            "SUBSTITUTE" => {
                need!(3);
                let t = arg_text!(0);
                let from = arg_text!(1);
                let to = arg_text!(2);
                Value::Text(t.replace(&from, &to))
            }
            "VALUE" => {
                need!(1);
                match arg_text!(0).trim().parse::<f64>() {
                    Ok(n) => Value::Number(n),
                    Err(_) => Value::Err("#VALUE!".into()),
                }
            }
            "SUMIF" | "COUNTIF" | "AVERAGEIF" => {
                need!(2);
                let (w, h, crit_vals) = match self.range_grid(sheet, &args[0]) {
                    Ok(g) => g,
                    Err(e) => return e,
                };
                let criteria = self.eval(sheet, &args[1]);
                if let Value::Err(_) = criteria {
                    return criteria;
                }
                let sum_vals = if args.len() > 2 {
                    match self.range_grid(sheet, &args[2]) {
                        Ok((w2, h2, v)) => {
                            if (w2, h2) != (w, h) {
                                return Value::Err("#VALUE! (ranges differ in size)".into());
                            }
                            Some(v)
                        }
                        Err(e) => return e,
                    }
                } else {
                    None
                };
                let mut total = 0.0;
                let mut count = 0usize;
                for (i, v) in crit_vals.iter().enumerate() {
                    if !matches_criteria(v, &criteria) {
                        continue;
                    }
                    count += 1;
                    let target = sum_vals.as_ref().map(|s| &s[i]).unwrap_or(v);
                    if let Value::Number(n) = target {
                        total += n;
                    }
                }
                match name {
                    "COUNTIF" => Value::Number(count as f64),
                    "AVERAGEIF" => {
                        if count == 0 {
                            Value::Err("#DIV/0!".into())
                        } else {
                            Value::Number(total / count as f64)
                        }
                    }
                    _ => Value::Number(total),
                }
            }
            "VLOOKUP" => {
                need!(3);
                let lookup = self.eval(sheet, &args[0]);
                let (w, h, grid) = match self.range_grid(sheet, &args[1]) {
                    Ok(g) => g,
                    Err(e) => return e,
                };
                let col_idx = arg_num!(2) as usize;
                if col_idx < 1 || col_idx > w {
                    return Value::Err("#REF!".into());
                }
                let exact = if args.len() > 3 {
                    match self.eval(sheet, &args[3]).as_bool() {
                        Ok(b) => !b,
                        Err(e) => return e,
                    }
                } else {
                    false
                };
                let mut best: Option<usize> = None;
                for r in 0..h {
                    let key = &grid[r * w];
                    let ord = compare_values(key, &lookup);
                    if ord == std::cmp::Ordering::Equal {
                        best = Some(r);
                        break;
                    }
                    if !exact && ord == std::cmp::Ordering::Less {
                        best = Some(r);
                    }
                }
                match best {
                    Some(r) => grid[r * w + col_idx - 1].clone(),
                    None => Value::Err("#N/A".into()),
                }
            }
            "INDEX" => {
                need!(2);
                let (w, h, grid) = match self.range_grid(sheet, &args[0]) {
                    Ok(g) => g,
                    Err(e) => return e,
                };
                let r = arg_num!(1) as usize;
                let c = if args.len() > 2 { arg_num!(2) as usize } else { 1 };
                if r < 1 || r > h || c < 1 || c > w {
                    return Value::Err("#REF!".into());
                }
                grid[(r - 1) * w + (c - 1)].clone()
            }
            "MATCH" => {
                need!(2);
                let lookup = self.eval(sheet, &args[0]);
                let (w, h, grid) = match self.range_grid(sheet, &args[1]) {
                    Ok(g) => g,
                    Err(e) => return e,
                };
                if w != 1 && h != 1 {
                    return Value::Err("#N/A (MATCH needs a vector)".into());
                }
                let mode = if args.len() > 2 { arg_num!(2) } else { 1.0 };
                let mut best: Option<usize> = None;
                for (i, v) in grid.iter().enumerate() {
                    let ord = compare_values(v, &lookup);
                    if ord == std::cmp::Ordering::Equal {
                        best = Some(i);
                        if mode == 0.0 {
                            break;
                        }
                    } else if mode > 0.0 && ord == std::cmp::Ordering::Less {
                        best = Some(i);
                    }
                }
                match best {
                    Some(i) => Value::Number((i + 1) as f64),
                    None => Value::Err("#N/A".into()),
                }
            }
            "TODAY" | "NOW" => {
                let secs = std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .map(|d| d.as_secs_f64())
                    .unwrap_or(0.0);
                let serial = secs / 86_400.0 + 25_569.0;
                Value::Number(if name == "TODAY" { serial.floor() } else { serial })
            }
            "DATE" => {
                need!(3);
                let (y, m, d) = (arg_num!(0) as i64, arg_num!(1) as u32, arg_num!(2) as u32);
                if !(1..=12).contains(&m) || !(1..=31).contains(&d) {
                    return Value::Err("#NUM!".into());
                }
                Value::Number(crate::xlsx::date_serial(y, m, d))
            }
            "YEAR" | "MONTH" | "DAY" => {
                need!(1);
                let serial = arg_num!(0);
                let (y, m, d) = crate::xlsx::serial_civil(serial);
                Value::Number(match name {
                    "YEAR" => y as f64,
                    "MONTH" => m as f64,
                    _ => d as f64,
                })
            }
            other => Value::Err(format!("#NAME? (unsupported function {other})")),
        }
    }
}

trait MapEmpty {
    fn map_empty(self) -> Value;
}

impl MapEmpty for Value {
    /// MIN/MAX of an empty set is 0 in Excel.
    fn map_empty(self) -> Value {
        match self {
            Value::Number(n) if n.is_infinite() => Value::Number(0.0),
            other => other,
        }
    }
}

/// Excel-flavored comparison: numbers numerically, text case-insensitively,
/// mixed types by type rank (number < text < bool).
fn compare_values(a: &Value, b: &Value) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    fn rank(v: &Value) -> u8 {
        match v {
            Value::Number(_) | Value::Empty => 0,
            Value::Text(_) => 1,
            Value::Bool(_) => 2,
            Value::Err(_) => 3,
        }
    }
    match (a, b) {
        (Value::Number(x), Value::Number(y)) => x.partial_cmp(y).unwrap_or(Ordering::Equal),
        (Value::Empty, Value::Number(y)) => 0f64.partial_cmp(y).unwrap_or(Ordering::Equal),
        (Value::Number(x), Value::Empty) => x.partial_cmp(&0.0).unwrap_or(Ordering::Equal),
        (Value::Text(x), Value::Text(y)) => x.to_lowercase().cmp(&y.to_lowercase()),
        (Value::Empty, Value::Text(y)) => "".cmp(y.to_lowercase().as_str()),
        (Value::Text(x), Value::Empty) => x.to_lowercase().as_str().cmp(""),
        (Value::Bool(x), Value::Bool(y)) => x.cmp(y),
        _ => rank(a).cmp(&rank(b)),
    }
}

/// SUMIF/COUNTIF criteria: ">5", "<>x", "=y", or a plain value.
fn matches_criteria(v: &Value, criteria: &Value) -> bool {
    use std::cmp::Ordering;
    let crit_text = criteria.as_text();
    let (op, rest): (&str, &str) = if let Some(r) = crit_text.strip_prefix(">=") {
        (">=", r)
    } else if let Some(r) = crit_text.strip_prefix("<=") {
        ("<=", r)
    } else if let Some(r) = crit_text.strip_prefix("<>") {
        ("<>", r)
    } else if let Some(r) = crit_text.strip_prefix('>') {
        (">", r)
    } else if let Some(r) = crit_text.strip_prefix('<') {
        ("<", r)
    } else if let Some(r) = crit_text.strip_prefix('=') {
        ("=", r)
    } else {
        ("=", crit_text.as_str())
    };
    let target = match rest.trim().parse::<f64>() {
        Ok(n) => Value::Number(n),
        Err(_) => Value::Text(rest.to_string()),
    };
    // A plain (non-text) criteria value compares directly.
    let target = if op == "=" && !matches!(criteria, Value::Text(_)) {
        criteria.clone()
    } else {
        target
    };
    if matches!(v, Value::Empty) && op == "=" && rest.is_empty() {
        return true;
    }
    if matches!(v, Value::Empty) {
        return false;
    }
    let ord = compare_values(v, &target);
    match op {
        ">" => ord == Ordering::Greater,
        "<" => ord == Ordering::Less,
        ">=" => ord != Ordering::Less,
        "<=" => ord != Ordering::Greater,
        "<>" => ord != Ordering::Equal,
        _ => ord == Ordering::Equal,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;

    struct Grid(HashMap<(String, u32, u32), CellContent>);

    impl CellSource for Grid {
        fn cell(&self, sheet: &str, col: u32, row: u32) -> CellContent {
            self.0
                .get(&(sheet.to_string(), col, row))
                .cloned()
                .unwrap_or(CellContent::Empty)
        }
        fn has_sheet(&self, sheet: &str) -> bool {
            sheet == "S1" || sheet == "Other"
        }
    }

    fn grid() -> Grid {
        let mut m = HashMap::new();
        // S1!A1:A4 = 10, 20, 30, "x"; B1 = formula A1*2; B2 circular
        m.insert(("S1".into(), 0, 0), CellContent::Number(10.0));
        m.insert(("S1".into(), 0, 1), CellContent::Number(20.0));
        m.insert(("S1".into(), 0, 2), CellContent::Number(30.0));
        m.insert(("S1".into(), 0, 3), CellContent::Text("x".into()));
        m.insert(("S1".into(), 1, 0), CellContent::Formula("A1*2".into()));
        m.insert(("S1".into(), 1, 1), CellContent::Formula("B2+1".into()));
        m.insert(("Other".into(), 0, 0), CellContent::Number(7.0));
        Grid(m)
    }

    fn eval(f: &str) -> Value {
        let g = grid();
        Evaluator::new(&g).eval_formula("S1", f)
    }

    #[test]
    fn arithmetic_and_refs() {
        assert_eq!(eval("=1+2*3"), Value::Number(7.0));
        assert_eq!(eval("=(1+2)*3^2"), Value::Number(27.0));
        assert_eq!(eval("=-A1+5%"), Value::Number(-9.95));
        assert_eq!(eval("=SUM(A1:A4)"), Value::Number(60.0));
        assert_eq!(eval("=B1"), Value::Number(20.0)); // nested formula
        assert_eq!(eval("=Other!A1+1"), Value::Number(8.0));
        assert_eq!(eval("=$A$2*2"), Value::Number(40.0));
    }

    #[test]
    fn functions() {
        assert_eq!(eval("=AVERAGE(A1:A3)"), Value::Number(20.0));
        assert_eq!(eval("=IF(A1>5,\"big\",\"small\")"), Value::Text("big".into()));
        assert_eq!(eval("=COUNTIF(A1:A3,\">15\")"), Value::Number(2.0));
        assert_eq!(eval("=SUMIF(A1:A3,\">=20\")"), Value::Number(50.0));
        assert_eq!(eval("=CONCAT(\"a\",A1,\"b\")"), Value::Text("a10b".into()));
        assert_eq!(eval("=LEFT(\"hello\",2)"), Value::Text("he".into()));
        assert_eq!(eval("=VLOOKUP(20,A1:B3,1,FALSE)"), Value::Number(20.0));
        assert_eq!(eval("=MATCH(30,A1:A3,0)"), Value::Number(3.0));
        assert_eq!(eval("=INDEX(A1:A3,2)"), Value::Number(20.0));
        assert_eq!(eval("=IFERROR(1/0,\"oops\")"), Value::Text("oops".into()));
        assert_eq!(eval("=ROUND(2.345,2)"), Value::Number(2.35));
        assert_eq!(eval("=YEAR(DATE(2026,7,11))"), Value::Number(2026.0));
    }

    #[test]
    fn errors() {
        assert_eq!(eval("=1/0"), Value::Err("#DIV/0!".into()));
        assert_eq!(eval("=B2"), Value::Err("#CIRC!".into())); // self-referencing
        assert_eq!(eval("=NOSUCHFN(1)"), Value::Err("#NAME? (unsupported function NOSUCHFN)".into()));
        assert_eq!(eval("=Missing!A1"), Value::Err("#REF!".into()));
    }
}
