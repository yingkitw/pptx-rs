//! Shape creation support for PPTX generation
//!
//! Provides shape types, fills, lines, and builders for creating shapes in slides.

/// Shape types available in PPTX
#[derive(Clone, Debug, Copy, PartialEq)]
pub enum ShapeType {
    // Basic shapes
    Rectangle,
    RoundedRectangle,
    Ellipse,
    Circle, // Alias for Ellipse
    Triangle,
    RightTriangle,
    Diamond,
    Pentagon,
    Hexagon,
    Octagon,
    
    // Arrows
    RightArrow,
    LeftArrow,
    UpArrow,
    DownArrow,
    LeftRightArrow,
    UpDownArrow,
    BentArrow,
    UTurnArrow,
    
    // Stars and banners
    Star4,
    Star5,
    Star6,
    Star8,
    Ribbon,
    Wave,
    
    // Callouts
    WedgeRectCallout,
    WedgeEllipseCallout,
    CloudCallout,
    
    // Flow chart
    FlowChartProcess,
    FlowChartDecision,
    FlowChartTerminator,
    FlowChartDocument,
    
    // Other
    Heart,
    Lightning,
    Sun,
    Moon,
    Cloud,
    Brace,
    Bracket,
    Plus,
    Minus,
}

impl ShapeType {
    /// Get the preset geometry name for the shape (OOXML preset name)
    pub fn preset_name(&self) -> &'static str {
        match self {
            ShapeType::Rectangle => "rect",
            ShapeType::RoundedRectangle => "roundRect",
            ShapeType::Ellipse | ShapeType::Circle => "ellipse",
            ShapeType::Triangle => "triangle",
            ShapeType::RightTriangle => "rtTriangle",
            ShapeType::Diamond => "diamond",
            ShapeType::Pentagon => "pentagon",
            ShapeType::Hexagon => "hexagon",
            ShapeType::Octagon => "octagon",
            
            ShapeType::RightArrow => "rightArrow",
            ShapeType::LeftArrow => "leftArrow",
            ShapeType::UpArrow => "upArrow",
            ShapeType::DownArrow => "downArrow",
            ShapeType::LeftRightArrow => "leftRightArrow",
            ShapeType::UpDownArrow => "upDownArrow",
            ShapeType::BentArrow => "bentArrow",
            ShapeType::UTurnArrow => "uturnArrow",
            
            ShapeType::Star4 => "star4",
            ShapeType::Star5 => "star5",
            ShapeType::Star6 => "star6",
            ShapeType::Star8 => "star8",
            ShapeType::Ribbon => "ribbon2",
            ShapeType::Wave => "wave",
            
            ShapeType::WedgeRectCallout => "wedgeRectCallout",
            ShapeType::WedgeEllipseCallout => "wedgeEllipseCallout",
            ShapeType::CloudCallout => "cloudCallout",
            
            ShapeType::FlowChartProcess => "flowChartProcess",
            ShapeType::FlowChartDecision => "flowChartDecision",
            ShapeType::FlowChartTerminator => "flowChartTerminator",
            ShapeType::FlowChartDocument => "flowChartDocument",
            
            ShapeType::Heart => "heart",
            ShapeType::Lightning => "lightningBolt",
            ShapeType::Sun => "sun",
            ShapeType::Moon => "moon",
            ShapeType::Cloud => "cloud",
            ShapeType::Brace => "leftBrace",
            ShapeType::Bracket => "leftBracket",
            ShapeType::Plus => "mathPlus",
            ShapeType::Minus => "mathMinus",
        }
    }

    /// Get a user-friendly name for the shape
    pub fn display_name(&self) -> &'static str {
        match self {
            ShapeType::Rectangle => "Rectangle",
            ShapeType::RoundedRectangle => "Rounded Rectangle",
            ShapeType::Ellipse => "Ellipse",
            ShapeType::Circle => "Circle",
            ShapeType::Triangle => "Triangle",
            ShapeType::RightTriangle => "Right Triangle",
            ShapeType::Diamond => "Diamond",
            ShapeType::Pentagon => "Pentagon",
            ShapeType::Hexagon => "Hexagon",
            ShapeType::Octagon => "Octagon",
            ShapeType::RightArrow => "Right Arrow",
            ShapeType::LeftArrow => "Left Arrow",
            ShapeType::UpArrow => "Up Arrow",
            ShapeType::DownArrow => "Down Arrow",
            ShapeType::LeftRightArrow => "Left-Right Arrow",
            ShapeType::UpDownArrow => "Up-Down Arrow",
            ShapeType::BentArrow => "Bent Arrow",
            ShapeType::UTurnArrow => "U-Turn Arrow",
            ShapeType::Star4 => "4-Point Star",
            ShapeType::Star5 => "5-Point Star",
            ShapeType::Star6 => "6-Point Star",
            ShapeType::Star8 => "8-Point Star",
            ShapeType::Ribbon => "Ribbon",
            ShapeType::Wave => "Wave",
            ShapeType::WedgeRectCallout => "Rectangle Callout",
            ShapeType::WedgeEllipseCallout => "Oval Callout",
            ShapeType::CloudCallout => "Cloud Callout",
            ShapeType::FlowChartProcess => "Process",
            ShapeType::FlowChartDecision => "Decision",
            ShapeType::FlowChartTerminator => "Terminator",
            ShapeType::FlowChartDocument => "Document",
            ShapeType::Heart => "Heart",
            ShapeType::Lightning => "Lightning Bolt",
            ShapeType::Sun => "Sun",
            ShapeType::Moon => "Moon",
            ShapeType::Cloud => "Cloud",
            ShapeType::Brace => "Brace",
            ShapeType::Bracket => "Bracket",
            ShapeType::Plus => "Plus",
            ShapeType::Minus => "Minus",
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
        assert_eq!(ShapeType::RightArrow.preset_name(), "rightArrow");
        assert_eq!(ShapeType::Star5.preset_name(), "star5");
        assert_eq!(ShapeType::Heart.preset_name(), "heart");
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
