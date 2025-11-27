//! Shape XML generation for PPTX
//!
//! Generates XML for shapes embedded in slides.

use super::shapes::{Shape, ShapeFill, ShapeLine};

/// Escape XML special characters
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Generate XML for a shape
pub fn generate_shape_xml(shape: &Shape, shape_id: u32) -> String {
    let fill_xml = generate_fill_xml(&shape.fill);
    let line_xml = generate_line_xml(&shape.line);
    let text_xml = generate_text_xml(&shape.text);
    
    format!(
        r#"<p:sp>
<p:nvSpPr>
<p:cNvPr id="{}" name="Shape {}"/>
<p:cNvSpPr/>
<p:nvPr/>
</p:nvSpPr>
<p:spPr>
<a:xfrm>
<a:off x="{}" y="{}"/>
<a:ext cx="{}" cy="{}"/>
</a:xfrm>
<a:prstGeom prst="{}">
<a:avLst/>
</a:prstGeom>
{}{}
</p:spPr>
{}
</p:sp>"#,
        shape_id,
        shape_id,
        shape.x,
        shape.y,
        shape.width,
        shape.height,
        shape.shape_type.preset_name(),
        fill_xml,
        line_xml,
        text_xml,
    )
}

/// Generate fill XML
fn generate_fill_xml(fill: &Option<ShapeFill>) -> String {
    match fill {
        Some(f) => {
            let alpha = f.transparency
                .map(|t| format!(r#"<a:alpha val="{}"/>"#, t))
                .unwrap_or_default();
            
            format!(
                r#"<a:solidFill>
<a:srgbClr val="{}">{}</a:srgbClr>
</a:solidFill>"#,
                f.color, alpha
            )
        }
        None => String::new(),
    }
}

/// Generate line XML
fn generate_line_xml(line: &Option<ShapeLine>) -> String {
    match line {
        Some(l) => {
            format!(
                r#"<a:ln w="{}">
<a:solidFill>
<a:srgbClr val="{}"/>
</a:solidFill>
</a:ln>"#,
                l.width, l.color
            )
        }
        None => String::new(),
    }
}

/// Generate text body XML for shape
fn generate_text_xml(text: &Option<String>) -> String {
    match text {
        Some(t) => {
            format!(
                r#"<p:txBody>
<a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/>
<a:lstStyle/>
<a:p>
<a:pPr algn="ctr"/>
<a:r>
<a:rPr lang="en-US" sz="1800" dirty="0"/>
<a:t>{}</a:t>
</a:r>
</a:p>
</p:txBody>"#,
                escape_xml(t)
            )
        }
        None => {
            // Empty text body required for shapes
            r#"<p:txBody>
<a:bodyPr/>
<a:lstStyle/>
<a:p/>
</p:txBody>"#.to_string()
        }
    }
}

/// Generate XML for multiple shapes
pub fn generate_shapes_xml(shapes: &[Shape], start_id: u32) -> String {
    shapes.iter()
        .enumerate()
        .map(|(i, shape)| generate_shape_xml(shape, start_id + i as u32))
        .collect::<Vec<_>>()
        .join("\n")
}

/// Generate connector shape XML (for arrows connecting shapes)
pub fn generate_connector_xml(
    start_x: u32, start_y: u32,
    end_x: u32, end_y: u32,
    shape_id: u32,
    color: &str,
    width: u32,
) -> String {
    format!(
        r#"<p:cxnSp>
<p:nvCxnSpPr>
<p:cNvPr id="{}" name="Connector {}"/>
<p:cNvCxnSpPr/>
<p:nvPr/>
</p:nvCxnSpPr>
<p:spPr>
<a:xfrm>
<a:off x="{}" y="{}"/>
<a:ext cx="{}" cy="{}"/>
</a:xfrm>
<a:prstGeom prst="straightConnector1">
<a:avLst/>
</a:prstGeom>
<a:ln w="{}">
<a:solidFill>
<a:srgbClr val="{}"/>
</a:solidFill>
<a:tailEnd type="triangle"/>
</a:ln>
</p:spPr>
</p:cxnSp>"#,
        shape_id,
        shape_id,
        start_x.min(end_x),
        start_y.min(end_y),
        (end_x as i64 - start_x as i64).unsigned_abs() as u32,
        (end_y as i64 - start_y as i64).unsigned_abs() as u32,
        width,
        color.trim_start_matches('#').to_uppercase(),
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use super::super::shapes::ShapeType;

    #[test]
    fn test_generate_shape_xml() {
        let shape = Shape::new(ShapeType::Rectangle, 100000, 200000, 500000, 300000)
            .with_fill(ShapeFill::new("FF0000"));
        
        let xml = generate_shape_xml(&shape, 10);
        
        assert!(xml.contains("p:sp"));
        assert!(xml.contains("id=\"10\""));
        assert!(xml.contains("rect"));
        assert!(xml.contains("FF0000"));
    }

    #[test]
    fn test_generate_shape_with_text() {
        let shape = Shape::new(ShapeType::Rectangle, 0, 0, 1000000, 500000)
            .with_text("Hello World");
        
        let xml = generate_shape_xml(&shape, 1);
        
        assert!(xml.contains("Hello World"));
        assert!(xml.contains("p:txBody"));
    }

    #[test]
    fn test_generate_shape_with_line() {
        let shape = Shape::new(ShapeType::Circle, 0, 0, 500000, 500000)
            .with_line(ShapeLine::new("000000", 25400));
        
        let xml = generate_shape_xml(&shape, 1);
        
        assert!(xml.contains("a:ln"));
        assert!(xml.contains("25400"));
    }

    #[test]
    fn test_generate_multiple_shapes() {
        let shapes = vec![
            Shape::new(ShapeType::Rectangle, 0, 0, 100000, 100000),
            Shape::new(ShapeType::Circle, 200000, 0, 100000, 100000),
        ];
        
        let xml = generate_shapes_xml(&shapes, 10);
        
        assert!(xml.contains("id=\"10\""));
        assert!(xml.contains("id=\"11\""));
        assert!(xml.contains("rect"));
        assert!(xml.contains("ellipse"));
    }

    #[test]
    fn test_generate_connector() {
        let xml = generate_connector_xml(0, 0, 1000000, 500000, 1, "0000FF", 12700);
        
        assert!(xml.contains("p:cxnSp"));
        assert!(xml.contains("straightConnector1"));
        assert!(xml.contains("triangle")); // arrow head
    }

    #[test]
    fn test_escape_xml_in_text() {
        let shape = Shape::new(ShapeType::Rectangle, 0, 0, 100000, 100000)
            .with_text("A < B & C > D");
        
        let xml = generate_shape_xml(&shape, 1);
        
        assert!(xml.contains("A &lt; B &amp; C &gt; D"));
    }
}
