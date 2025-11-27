//! Drawing Markup Language (DML) XML elements
//!
//! Core DrawingML types used across OOXML documents.

use super::xmlchemy::XmlElement;

/// Color types in DrawingML
#[derive(Debug, Clone)]
pub enum Color {
    /// RGB color (e.g., "FF0000" for red)
    Rgb(String),
    /// Scheme color (e.g., "accent1", "dk1")
    Scheme(String),
    /// System color (e.g., "windowText")
    System(String),
}

impl Color {
    pub fn rgb(hex: &str) -> Self {
        Color::Rgb(hex.trim_start_matches('#').to_uppercase())
    }

    pub fn scheme(name: &str) -> Self {
        Color::Scheme(name.to_string())
    }

    pub fn parse(elem: &XmlElement) -> Option<Self> {
        if let Some(srgb) = elem.find("srgbClr") {
            return srgb.attr("val").map(|v| Color::Rgb(v.to_string()));
        }
        if let Some(scheme) = elem.find("schemeClr") {
            return scheme.attr("val").map(|v| Color::Scheme(v.to_string()));
        }
        if let Some(sys) = elem.find("sysClr") {
            return sys.attr("val").map(|v| Color::System(v.to_string()));
        }
        None
    }

    pub fn to_xml(&self) -> String {
        match self {
            Color::Rgb(hex) => format!(r#"<a:srgbClr val="{}"/>"#, hex),
            Color::Scheme(name) => format!(r#"<a:schemeClr val="{}"/>"#, name),
            Color::System(name) => format!(r#"<a:sysClr val="{}"/>"#, name),
        }
    }
}

/// Effect extent (a:effectExtent)
#[derive(Debug, Clone, Default)]
pub struct EffectExtent {
    pub left: i64,
    pub top: i64,
    pub right: i64,
    pub bottom: i64,
}

impl EffectExtent {
    pub fn parse(elem: &XmlElement) -> Self {
        EffectExtent {
            left: elem.attr("l").and_then(|v| v.parse().ok()).unwrap_or(0),
            top: elem.attr("t").and_then(|v| v.parse().ok()).unwrap_or(0),
            right: elem.attr("r").and_then(|v| v.parse().ok()).unwrap_or(0),
            bottom: elem.attr("b").and_then(|v| v.parse().ok()).unwrap_or(0),
        }
    }

    pub fn to_xml(&self) -> String {
        format!(
            r#"<a:effectExtent l="{}" t="{}" r="{}" b="{}"/>"#,
            self.left, self.top, self.right, self.bottom
        )
    }
}

/// Outline (a:ln) - line/border properties
#[derive(Debug, Clone, Default)]
pub struct Outline {
    pub width: Option<u32>,
    pub cap: Option<String>,
    pub compound: Option<String>,
    pub color: Option<Color>,
    pub dash: Option<String>,
}

impl Outline {
    pub fn new() -> Self {
        Outline::default()
    }

    pub fn with_width(mut self, width: u32) -> Self {
        self.width = Some(width);
        self
    }

    pub fn with_color(mut self, color: Color) -> Self {
        self.color = Some(color);
        self
    }

    pub fn parse(elem: &XmlElement) -> Self {
        let mut outline = Outline::new();
        
        outline.width = elem.attr("w").and_then(|v| v.parse().ok());
        outline.cap = elem.attr("cap").map(|s| s.to_string());
        outline.compound = elem.attr("cmpd").map(|s| s.to_string());

        if let Some(solid_fill) = elem.find("solidFill") {
            outline.color = Color::parse(solid_fill);
        }

        if let Some(prst_dash) = elem.find("prstDash") {
            outline.dash = prst_dash.attr("val").map(|s| s.to_string());
        }

        outline
    }

    pub fn to_xml(&self) -> String {
        let width = self.width.unwrap_or(12700);
        let mut inner = String::new();

        if let Some(ref color) = self.color {
            inner.push_str("<a:solidFill>");
            inner.push_str(&color.to_xml());
            inner.push_str("</a:solidFill>");
        }

        if let Some(ref dash) = self.dash {
            inner.push_str(&format!(r#"<a:prstDash val="{}"/>"#, dash));
        }

        if inner.is_empty() {
            format!(r#"<a:ln w="{}"/>"#, width)
        } else {
            format!(r#"<a:ln w="{}">{}</a:ln>"#, width, inner)
        }
    }
}

/// Gradient stop
#[derive(Debug, Clone)]
pub struct GradientStop {
    pub position: u32, // 0-100000 (percentage * 1000)
    pub color: Color,
}

impl GradientStop {
    pub fn new(position: u32, color: Color) -> Self {
        GradientStop { position, color }
    }

    pub fn to_xml(&self) -> String {
        format!(
            r#"<a:gs pos="{}">{}</a:gs>"#,
            self.position,
            self.color.to_xml()
        )
    }
}

/// Gradient fill
#[derive(Debug, Clone)]
pub struct GradientFill {
    pub stops: Vec<GradientStop>,
    pub angle: Option<i32>, // in 60000ths of a degree
}

impl GradientFill {
    pub fn new() -> Self {
        GradientFill {
            stops: Vec::new(),
            angle: None,
        }
    }

    pub fn add_stop(mut self, position: u32, color: Color) -> Self {
        self.stops.push(GradientStop::new(position, color));
        self
    }

    pub fn with_angle(mut self, degrees: i32) -> Self {
        self.angle = Some(degrees * 60000);
        self
    }

    pub fn to_xml(&self) -> String {
        let mut xml = String::from("<a:gradFill><a:gsLst>");
        for stop in &self.stops {
            xml.push_str(&stop.to_xml());
        }
        xml.push_str("</a:gsLst>");

        if let Some(angle) = self.angle {
            xml.push_str(&format!(r#"<a:lin ang="{}" scaled="1"/>"#, angle));
        }

        xml.push_str("</a:gradFill>");
        xml
    }
}

impl Default for GradientFill {
    fn default() -> Self {
        Self::new()
    }
}

/// Pattern fill type
#[derive(Debug, Clone)]
pub struct PatternFill {
    pub preset: String,
    pub foreground: Color,
    pub background: Color,
}

impl PatternFill {
    pub fn new(preset: &str, fg: Color, bg: Color) -> Self {
        PatternFill {
            preset: preset.to_string(),
            foreground: fg,
            background: bg,
        }
    }

    pub fn to_xml(&self) -> String {
        format!(
            r#"<a:pattFill prst="{}"><a:fgClr>{}</a:fgClr><a:bgClr>{}</a:bgClr></a:pattFill>"#,
            self.preset,
            self.foreground.to_xml(),
            self.background.to_xml()
        )
    }
}

/// Fill types
#[derive(Debug, Clone)]
pub enum Fill {
    None,
    Solid(Color),
    Gradient(GradientFill),
    Pattern(PatternFill),
}

impl Fill {
    pub fn solid(color: Color) -> Self {
        Fill::Solid(color)
    }

    pub fn to_xml(&self) -> String {
        match self {
            Fill::None => "<a:noFill/>".to_string(),
            Fill::Solid(color) => format!("<a:solidFill>{}</a:solidFill>", color.to_xml()),
            Fill::Gradient(grad) => grad.to_xml(),
            Fill::Pattern(pat) => pat.to_xml(),
        }
    }
}

/// Point in EMUs
#[derive(Debug, Clone, Copy, Default)]
pub struct Point {
    pub x: i64,
    pub y: i64,
}

impl Point {
    pub fn new(x: i64, y: i64) -> Self {
        Point { x, y }
    }

    pub fn from_inches(x: f64, y: f64) -> Self {
        Point {
            x: (x * 914400.0) as i64,
            y: (y * 914400.0) as i64,
        }
    }
}

/// Size in EMUs
#[derive(Debug, Clone, Copy, Default)]
pub struct Size {
    pub width: i64,
    pub height: i64,
}

impl Size {
    pub fn new(width: i64, height: i64) -> Self {
        Size { width, height }
    }

    pub fn from_inches(width: f64, height: f64) -> Self {
        Size {
            width: (width * 914400.0) as i64,
            height: (height * 914400.0) as i64,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_color_rgb() {
        let color = Color::rgb("FF0000");
        let xml = color.to_xml();
        assert!(xml.contains("srgbClr"));
        assert!(xml.contains("FF0000"));
    }

    #[test]
    fn test_color_scheme() {
        let color = Color::scheme("accent1");
        let xml = color.to_xml();
        assert!(xml.contains("schemeClr"));
        assert!(xml.contains("accent1"));
    }

    #[test]
    fn test_outline_to_xml() {
        let outline = Outline::new()
            .with_width(25400)
            .with_color(Color::rgb("0000FF"));
        let xml = outline.to_xml();
        
        assert!(xml.contains("w=\"25400\""));
        assert!(xml.contains("0000FF"));
    }

    #[test]
    fn test_gradient_fill() {
        let grad = GradientFill::new()
            .add_stop(0, Color::rgb("FF0000"))
            .add_stop(100000, Color::rgb("0000FF"))
            .with_angle(90);
        
        let xml = grad.to_xml();
        assert!(xml.contains("gradFill"));
        assert!(xml.contains("FF0000"));
        assert!(xml.contains("0000FF"));
    }

    #[test]
    fn test_fill_solid() {
        let fill = Fill::solid(Color::rgb("00FF00"));
        let xml = fill.to_xml();
        assert!(xml.contains("solidFill"));
        assert!(xml.contains("00FF00"));
    }
}
