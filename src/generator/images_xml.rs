//! Image XML generation for PPTX presentations
//!
//! Generates proper PPTX XML for image embedding and display

use crate::generator::images::Image;

/// Generate image XML for a slide
pub fn generate_image_xml(image: &Image, shape_id: usize, rel_id: usize) -> String {
    let rel_id_str = format!("rId{}", rel_id);
    
    let xml = format!(
        r#"<p:pic>
<p:nvPicPr>
<p:cNvPr id="{}" name="{}"/>
<p:cNvPicPr>
<a:picLocks noChangeAspect="1"/>
</p:cNvPicPr>
<p:nvPr/>
</p:nvPicPr>
<p:blipFill>
<a:blip r:embed="{}"/>
<a:stretch>
<a:fillRect/>
</a:stretch>
</p:blipFill>
<p:spPr>
<a:xfrm>
<a:off x="{}" y="{}"/>
<a:ext cx="{}" cy="{}"/>
</a:xfrm>
<a:prstGeom prst="rect">
<a:avLst/>
</a:prstGeom>
</p:spPr>
</p:pic>"#,
        shape_id,
        escape_xml(&image.filename),
        rel_id_str,
        image.x,
        image.y,
        image.width,
        image.height
    );

    xml
}

/// Generate image relationship XML
pub fn generate_image_relationship(rel_id: usize, image_path: &str) -> String {
    format!(
        r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{}"/>"#,
        rel_id,
        escape_xml(image_path)
    )
}

/// Generate image content type entry
pub fn generate_image_content_type(extension: &str) -> String {
    let mime_type = match extension.to_lowercase().as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tiff" => "image/tiff",
        "svg" => "image/svg+xml",
        _ => "application/octet-stream",
    };

    format!(
        r#"<Default Extension="{}" ContentType="{}"/>"#,
        extension.to_lowercase(),
        mime_type
    )
}

/// Escape XML special characters
fn escape_xml(s: &str) -> String {
    s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\"", "&quot;")
        .replace("'", "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::generator::images::Image;

    #[test]
    fn test_generate_simple_image_xml() {
        let img = Image::new("photo.png", 1920000, 1080000, "PNG");
        let xml = generate_image_xml(&img, 1, 1);

        assert!(xml.contains("p:pic"));
        assert!(xml.contains("photo.png"));
        assert!(xml.contains("rId1"));
    }

    #[test]
    fn test_generate_image_with_position() {
        let img = Image::new("photo.png", 1920000, 1080000, "PNG")
            .position(500000, 1000000);
        let xml = generate_image_xml(&img, 1, 1);

        assert!(xml.contains("x=\"500000\""));
        assert!(xml.contains("y=\"1000000\""));
    }

    #[test]
    fn test_generate_image_with_dimensions() {
        let img = Image::new("photo.png", 1920000, 1080000, "PNG");
        let xml = generate_image_xml(&img, 1, 1);

        assert!(xml.contains("cx=\"1920000\""));
        assert!(xml.contains("cy=\"1080000\""));
    }

    #[test]
    fn test_generate_image_relationship() {
        let rel = generate_image_relationship(1, "../media/image1.png");
        assert!(rel.contains("rId1"));
        assert!(rel.contains("../media/image1.png"));
    }

    #[test]
    fn test_generate_image_content_type_png() {
        let ct = generate_image_content_type("png");
        assert!(ct.contains("image/png"));
    }

    #[test]
    fn test_generate_image_content_type_jpg() {
        let ct = generate_image_content_type("jpg");
        assert!(ct.contains("image/jpeg"));
    }

    #[test]
    fn test_generate_image_content_type_gif() {
        let ct = generate_image_content_type("gif");
        assert!(ct.contains("image/gif"));
    }

    #[test]
    fn test_escape_xml_in_filename() {
        let img = Image::new("photo & <test>.png", 100, 100, "PNG");
        let xml = generate_image_xml(&img, 1, 1);

        assert!(xml.contains("&amp;"));
        assert!(xml.contains("&lt;"));
        assert!(xml.contains("&gt;"));
    }

    #[test]
    fn test_generate_image_xml_structure() {
        let img = Image::new("photo.png", 1920000, 1080000, "PNG");
        let xml = generate_image_xml(&img, 1, 1);

        assert!(xml.contains("p:nvPicPr"));
        assert!(xml.contains("p:cNvPicPr"));
        assert!(xml.contains("a:picLocks"));
        assert!(xml.contains("p:blipFill"));
        assert!(xml.contains("a:blip"));
        assert!(xml.contains("a:stretch"));
        assert!(xml.contains("p:spPr"));
        assert!(xml.contains("a:xfrm"));
        assert!(xml.contains("a:prstGeom"));
    }
}
