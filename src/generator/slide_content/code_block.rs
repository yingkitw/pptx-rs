//! Code block types for syntax highlighting

/// A code block with syntax highlighting info
#[derive(Clone, Debug)]
pub struct CodeBlock {
    pub code: String,
    pub language: String,
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
}

impl CodeBlock {
    pub fn new(code: &str, language: &str) -> Self {
        Self {
            code: code.to_string(),
            language: language.to_string(),
            x: 500000,
            y: 1800000,
            width: 8000000,
            height: 4000000,
        }
    }
    
    pub fn position(mut self, x: i64, y: i64) -> Self {
        self.x = x;
        self.y = y;
        self
    }
    
    pub fn size(mut self, width: i64, height: i64) -> Self {
        self.width = width;
        self.height = height;
        self
    }
}

