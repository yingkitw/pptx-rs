//! Shape XML elements for OOXML
//!
//! Provides types for parsing and generating DrawingML shape elements.

use super::xmlchemy::XmlElement;

/// Transform properties (a:xfrm)
#[derive(Debug, Clone, Default)]
pub struct Transform2D {
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
    pub rotation: Option<i32>,
    pub flip_h: bool,
    pub flip_v: bool,
}

impl Transform2D {
    pub fn new(x: i64, y: i64, width: i64, height: i64) -> Self {
        Transform2D {
            x, y, width, height,
            rotation: None,
            flip_h: false,
            flip_v: false,
        }
    }

    pub fn parse(elem: &XmlElement) -> Self {
        let mut xfrm = Transform2D::default();

        if let Some(off) = elem.find("off") {
            xfrm.x = off.attr("x").and_then(|v| v.parse().ok()).unwrap_or(0);
            xfrm.y = off.attr("y").and_then(|v| v.parse().ok()).unwrap_or(0);
        }

        if let Some(ext) = elem.find("ext") {
            xfrm.width = ext.attr("cx").and_then(|v| v.parse().ok()).unwrap_or(0);
            xfrm.height = ext.attr("cy").and_then(|v| v.parse().ok()).unwrap_or(0);
        }

        xfrm.rotation = elem.attr("rot").and_then(|v| v.parse().ok());
        xfrm.flip_h = elem.attr("flipH").map(|v| v == "1").unwrap_or(false);
        xfrm.flip_v = elem.attr("flipV").map(|v| v == "1").unwrap_or(false);

        xfrm
    }

    pub fn to_xml(&self) -> String {
        let mut attrs = Vec::new();
        
        if let Some(rot) = self.rotation {
            attrs.push(format!(r#"rot="{}""#, rot));
        }
        if self.flip_h {
            attrs.push(r#"flipH="1""#.to_string());
        }
        if self.flip_v {
            attrs.push(r#"flipV="1""#.to_string());
        }

        let attr_str = if attrs.is_empty() { String::new() } else { format!(" {}", attrs.join(" ")) };

        format!(
            r#"<a:xfrm{}><a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/></a:xfrm>"#,
            attr_str, self.x, self.y, self.width, self.height
        )
    }
}

/// Preset geometry (a:prstGeom)
#[derive(Debug, Clone)]
pub struct PresetGeometry {
    pub preset: String,
}

impl PresetGeometry {
    pub fn new(preset: &str) -> Self {
        PresetGeometry { preset: preset.to_string() }
    }

    pub fn parse(elem: &XmlElement) -> Option<Self> {
        elem.attr("prst").map(|p| PresetGeometry::new(p))
    }

    pub fn to_xml(&self) -> String {
        format!(r#"<a:prstGeom prst="{}"><a:avLst/></a:prstGeom>"#, self.preset)
    }
}

/// Solid fill (a:solidFill)
#[derive(Debug, Clone)]
pub struct SolidFill {
    pub color: String,
}

impl SolidFill {
    pub fn new(color: &str) -> Self {
        SolidFill { color: color.trim_start_matches('#').to_uppercase() }
    }

    pub fn parse(elem: &XmlElement) -> Option<Self> {
        if let Some(srgb) = elem.find("srgbClr") {
            return srgb.attr("val").map(|c| SolidFill::new(c));
        }
        None
    }

    pub fn to_xml(&self) -> String {
        format!(r#"<a:solidFill><a:srgbClr val="{}"/></a:solidFill>"#, self.color)
    }
}

/// Line properties (a:ln)
#[derive(Debug, Clone, Default)]
pub struct LineProperties {
    pub width: Option<u32>,
    pub color: Option<String>,
    pub dash: Option<String>,
}

impl LineProperties {
    pub fn new() -> Self {
        LineProperties::default()
    }

    pub fn parse(elem: &XmlElement) -> Self {
        let mut line = LineProperties::new();
        
        line.width = elem.attr("w").and_then(|v| v.parse().ok());
        
        if let Some(solid_fill) = elem.find("solidFill") {
            if let Some(srgb) = solid_fill.find("srgbClr") {
                line.color = srgb.attr("val").map(|s| s.to_string());
            }
        }

        if let Some(prst_dash) = elem.find("prstDash") {
            line.dash = prst_dash.attr("val").map(|s| s.to_string());
        }

        line
    }

    pub fn to_xml(&self) -> String {
        let width = self.width.unwrap_or(12700);
        let mut inner = String::new();

        if let Some(ref color) = self.color {
            inner.push_str(&format!(r#"<a:solidFill><a:srgbClr val="{}"/></a:solidFill>"#, color));
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

/// Shape properties (p:spPr)
#[derive(Debug, Clone, Default)]
pub struct ShapeProperties {
    pub transform: Option<Transform2D>,
    pub geometry: Option<PresetGeometry>,
    pub fill: Option<SolidFill>,
    pub line: Option<LineProperties>,
}

impl ShapeProperties {
    pub fn new() -> Self {
        ShapeProperties::default()
    }

    pub fn parse(elem: &XmlElement) -> Self {
        let mut props = ShapeProperties::new();

        if let Some(xfrm) = elem.find("xfrm") {
            props.transform = Some(Transform2D::parse(xfrm));
        }

        if let Some(prst_geom) = elem.find("prstGeom") {
            props.geometry = PresetGeometry::parse(prst_geom);
        }

        if let Some(solid_fill) = elem.find("solidFill") {
            props.fill = SolidFill::parse(solid_fill);
        }

        if let Some(ln) = elem.find("ln") {
            props.line = Some(LineProperties::parse(ln));
        }

        props
    }

    pub fn to_xml(&self) -> String {
        let mut xml = String::from("<p:spPr>");

        if let Some(ref xfrm) = self.transform {
            xml.push_str(&xfrm.to_xml());
        }

        if let Some(ref geom) = self.geometry {
            xml.push_str(&geom.to_xml());
        }

        if let Some(ref fill) = self.fill {
            xml.push_str(&fill.to_xml());
        } else {
            xml.push_str("<a:noFill/>");
        }

        if let Some(ref line) = self.line {
            xml.push_str(&line.to_xml());
        }

        xml.push_str("</p:spPr>");
        xml
    }
}

/// Non-visual shape properties
#[derive(Debug, Clone)]
pub struct NonVisualProperties {
    pub id: u32,
    pub name: String,
}

impl NonVisualProperties {
    pub fn new(id: u32, name: &str) -> Self {
        NonVisualProperties { id, name: name.to_string() }
    }

    pub fn parse(elem: &XmlElement) -> Option<Self> {
        if let Some(cnv_pr) = elem.find("cNvPr") {
            let id = cnv_pr.attr("id").and_then(|v| v.parse().ok()).unwrap_or(0);
            let name = cnv_pr.attr("name").unwrap_or("Shape").to_string();
            return Some(NonVisualProperties { id, name });
        }
        None
    }

    pub fn to_xml(&self) -> String {
        format!(
            r#"<p:nvSpPr><p:cNvPr id="{}" name="{}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>"#,
            self.id, escape_xml(&self.name)
        )
    }
}

fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_transform_to_xml() {
        let xfrm = Transform2D::new(100, 200, 1000, 500);
        let xml = xfrm.to_xml();
        
        assert!(xml.contains("x=\"100\""));
        assert!(xml.contains("y=\"200\""));
        assert!(xml.contains("cx=\"1000\""));
        assert!(xml.contains("cy=\"500\""));
    }

    #[test]
    fn test_preset_geometry_to_xml() {
        let geom = PresetGeometry::new("rect");
        let xml = geom.to_xml();
        
        assert!(xml.contains("prst=\"rect\""));
    }

    #[test]
    fn test_solid_fill_to_xml() {
        let fill = SolidFill::new("FF0000");
        let xml = fill.to_xml();
        
        assert!(xml.contains("FF0000"));
    }

    #[test]
    fn test_line_properties_to_xml() {
        let mut line = LineProperties::new();
        line.width = Some(25400);
        line.color = Some("0000FF".to_string());
        
        let xml = line.to_xml();
        assert!(xml.contains("w=\"25400\""));
        assert!(xml.contains("0000FF"));
    }
}
