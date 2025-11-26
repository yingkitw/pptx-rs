//! Shape creation support for PPTX generation

/// Shape types
#[derive(Clone, Debug, Copy)]
pub enum ShapeType {
    Rectangle,
    Circle,
    Triangle,
    Diamond,
    Arrow,
    Star,
    Hexagon,
}

impl ShapeType {
    /// Get the preset geometry name for the shape
    pub fn preset_name(&self) -> &'static str {
        match self {
            ShapeType::Rectangle => "rect",
            ShapeType::Circle => "ellipse",
            ShapeType::Triangle => "triangle",
            ShapeType::Diamond => "diamond",
            ShapeType::Arrow => "rightArrow",
            ShapeType::Star => "star5",
            ShapeType::Hexagon => "hexagon",
        }
    }
}

/// Shape fill/color properties
#[derive(Clone, Debug)]
pub struct ShapeFill {
    pub color: String, // RGB hex color (e.g., "FF0000")
    pub transparency: Option<u32>, // 0-100000 (100000 = fully transparent)
}

impl ShapeFill {
    /// Create new shape fill with color
    pub fn new(color: &str) -> Self {
        ShapeFill {
            color: color.trim_start_matches('#').to_uppercase(),
            transparency: None,
        }
    }

    /// Set transparency (0-100 percent)
    pub fn transparency(mut self, percent: u32) -> Self {
        let alpha = ((100 - percent.min(100)) * 1000) as u32;
        self.transparency = Some(alpha);
        self
    }
}

/// Shape line/border properties
#[derive(Clone, Debug)]
pub struct ShapeLine {
    pub color: String,
    pub width: u32, // in EMU (English Metric Units)
}

impl ShapeLine {
    /// Create new shape line with color and width
    pub fn new(color: &str, width: u32) -> Self {
        ShapeLine {
            color: color.trim_start_matches('#').to_uppercase(),
            width,
        }
    }
}

/// Shape definition
#[derive(Clone, Debug)]
pub struct Shape {
    pub shape_type: ShapeType,
    pub x: u32,      // Position X in EMU
    pub y: u32,      // Position Y in EMU
    pub width: u32,  // Width in EMU
    pub height: u32, // Height in EMU
    pub fill: Option<ShapeFill>,
    pub line: Option<ShapeLine>,
    pub text: Option<String>,
}

impl Shape {
    /// Create a new shape
    pub fn new(shape_type: ShapeType, x: u32, y: u32, width: u32, height: u32) -> Self {
        Shape {
            shape_type,
            x,
            y,
            width,
            height,
            fill: None,
            line: None,
            text: None,
        }
    }

    /// Set shape fill
    pub fn with_fill(mut self, fill: ShapeFill) -> Self {
        self.fill = Some(fill);
        self
    }

    /// Set shape line
    pub fn with_line(mut self, line: ShapeLine) -> Self {
        self.line = Some(line);
        self
    }

    /// Set shape text
    pub fn with_text(mut self, text: &str) -> Self {
        self.text = Some(text.to_string());
        self
    }
}

/// Convert EMU (English Metric Units) to inches
pub fn emu_to_inches(emu: u32) -> f64 {
    emu as f64 / 914400.0
}

/// Convert inches to EMU
pub fn inches_to_emu(inches: f64) -> u32 {
    (inches * 914400.0) as u32
}

/// Convert centimeters to EMU
pub fn cm_to_emu(cm: f64) -> u32 {
    (cm * 360000.0) as u32
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_shape_type_names() {
        assert_eq!(ShapeType::Rectangle.preset_name(), "rect");
        assert_eq!(ShapeType::Circle.preset_name(), "ellipse");
        assert_eq!(ShapeType::Arrow.preset_name(), "rightArrow");
    }

    #[test]
    fn test_shape_fill_builder() {
        let fill = ShapeFill::new("FF0000").transparency(50);
        assert_eq!(fill.color, "FF0000");
        assert_eq!(fill.transparency, Some(50000));
    }

    #[test]
    fn test_shape_builder() {
        let shape = Shape::new(ShapeType::Rectangle, 0, 0, 1000000, 500000)
            .with_fill(ShapeFill::new("0000FF"))
            .with_line(ShapeLine::new("000000", 25400))
            .with_text("Hello");

        assert_eq!(shape.x, 0);
        assert_eq!(shape.width, 1000000);
        assert_eq!(shape.text, Some("Hello".to_string()));
    }

    #[test]
    fn test_emu_conversions() {
        let emu = inches_to_emu(1.0);
        assert_eq!(emu, 914400);
        assert!((emu_to_inches(emu) - 1.0).abs() < 0.001);
    }

    #[test]
    fn test_cm_to_emu() {
        let emu = cm_to_emu(2.54); // 1 inch
        assert_eq!(emu, 914400);
    }
}
