//! Tests for image XML generation
//!
//! Tests the image XML generation module

use ppt_rs::generator::{Image, ImageBuilder, generate_image_xml, generate_image_relationship, generate_image_content_type};

// ============================================================================
// IMAGE CREATION TESTS
// ============================================================================

#[test]
fn test_image_creation() {
    let img = Image::new("photo.png", 1920, 1080, "PNG");
    assert_eq!(img.filename, "photo.png");
    assert_eq!(img.width, 1920);
    assert_eq!(img.height, 1080);
    assert_eq!(img.format, "PNG");
}

#[test]
fn test_image_position() {
    let img = Image::new("photo.png", 1920, 1080, "PNG")
        .position(500000, 1000000);
    assert_eq!(img.x, 500000);
    assert_eq!(img.y, 1000000);
}

#[test]
fn test_image_aspect_ratio_16_9() {
    let img = Image::new("photo.png", 1920, 1080, "PNG");
    let ratio = img.aspect_ratio();
    assert!((ratio - 1.777).abs() < 0.01);
}

#[test]
fn test_image_aspect_ratio_4_3() {
    let img = Image::new("photo.png", 1024, 768, "PNG");
    let ratio = img.aspect_ratio();
    assert!((ratio - 1.333).abs() < 0.01);
}

#[test]
fn test_image_aspect_ratio_1_1() {
    let img = Image::new("photo.png", 1000, 1000, "PNG");
    let ratio = img.aspect_ratio();
    assert!((ratio - 1.0).abs() < 0.01);
}

#[test]
fn test_image_scale_to_width() {
    let img = Image::new("photo.png", 1920, 1080, "PNG")
        .scale_to_width(960);
    assert_eq!(img.width, 960);
    assert_eq!(img.height, 540);
}

#[test]
fn test_image_scale_to_height() {
    let img = Image::new("photo.png", 1920, 1080, "PNG")
        .scale_to_height(540);
    assert_eq!(img.width, 960);
    assert_eq!(img.height, 540);
}

#[test]
fn test_image_extension_png() {
    let img = Image::new("photo.png", 1920, 1080, "PNG");
    assert_eq!(img.extension(), "png");
}

#[test]
fn test_image_extension_jpg() {
    let img = Image::new("photo.jpg", 1920, 1080, "JPG");
    assert_eq!(img.extension(), "jpg");
}

#[test]
fn test_image_mime_type_png() {
    let img = Image::new("photo.png", 1920, 1080, "PNG");
    assert_eq!(img.mime_type(), "image/png");
}

#[test]
fn test_image_mime_type_jpg() {
    let img = Image::new("photo.jpg", 1920, 1080, "JPG");
    assert_eq!(img.mime_type(), "image/jpeg");
}

#[test]
fn test_image_mime_type_gif() {
    let img = Image::new("photo.gif", 1920, 1080, "GIF");
    assert_eq!(img.mime_type(), "image/gif");
}

#[test]
fn test_image_mime_type_bmp() {
    let img = Image::new("photo.bmp", 1920, 1080, "BMP");
    assert_eq!(img.mime_type(), "image/bmp");
}

// ============================================================================
// IMAGE BUILDER TESTS
// ============================================================================

#[test]
fn test_image_builder_basic() {
    let img = ImageBuilder::new("photo.png", 1920, 1080).build();
    assert_eq!(img.filename, "photo.png");
    assert_eq!(img.width, 1920);
    assert_eq!(img.height, 1080);
}

#[test]
fn test_image_builder_with_position() {
    let img = ImageBuilder::new("photo.png", 1920, 1080)
        .position(500000, 1000000)
        .build();
    assert_eq!(img.x, 500000);
    assert_eq!(img.y, 1000000);
}

#[test]
fn test_image_builder_with_format() {
    let img = ImageBuilder::new("photo.bin", 1920, 1080)
        .format("JPEG")
        .build();
    assert_eq!(img.format, "JPEG");
}

#[test]
fn test_image_builder_scale_to_width() {
    let img = ImageBuilder::new("photo.png", 1920, 1080)
        .scale_to_width(960)
        .build();
    assert_eq!(img.width, 960);
    assert_eq!(img.height, 540);
}

#[test]
fn test_image_builder_scale_to_height() {
    let img = ImageBuilder::new("photo.png", 1920, 1080)
        .scale_to_height(540)
        .build();
    assert_eq!(img.width, 960);
    assert_eq!(img.height, 540);
}

#[test]
fn test_image_builder_fluent() {
    let img = ImageBuilder::new("photo.png", 1920, 1080)
        .position(500000, 1000000)
        .scale_to_width(960)
        .build();
    
    assert_eq!(img.filename, "photo.png");
    assert_eq!(img.width, 960);
    assert_eq!(img.height, 540);
    assert_eq!(img.x, 500000);
    assert_eq!(img.y, 1000000);
}

#[test]
fn test_image_builder_auto_format_detection() {
    let img_png = ImageBuilder::new("photo.png", 100, 100).build();
    let img_jpg = ImageBuilder::new("photo.jpg", 100, 100).build();
    let img_gif = ImageBuilder::new("photo.gif", 100, 100).build();
    
    assert_eq!(img_png.format, "PNG");
    assert_eq!(img_jpg.format, "JPG");
    assert_eq!(img_gif.format, "GIF");
}

// ============================================================================
// IMAGE XML GENERATION TESTS
// ============================================================================

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
fn test_generate_image_with_different_shape_ids() {
    let img = Image::new("photo.png", 1920000, 1080000, "PNG");
    let xml1 = generate_image_xml(&img, 1, 1);
    let xml2 = generate_image_xml(&img, 2, 2);

    assert!(xml1.contains("id=\"1\""));
    assert!(xml2.contains("id=\"2\""));
}

#[test]
fn test_generate_image_relationship() {
    let rel = generate_image_relationship(1, "../media/image1.png");
    assert!(rel.contains("rId1"));
    assert!(rel.contains("../media/image1.png"));
}

#[test]
fn test_generate_image_relationship_multiple() {
    let rel1 = generate_image_relationship(1, "../media/image1.png");
    let rel2 = generate_image_relationship(2, "../media/image2.jpg");

    assert!(rel1.contains("rId1"));
    assert!(rel2.contains("rId2"));
}

#[test]
fn test_generate_image_content_type_png() {
    let ct = generate_image_content_type("png");
    assert!(ct.contains("png"));
    assert!(ct.contains("image/png"));
}

#[test]
fn test_generate_image_content_type_jpg() {
    let ct = generate_image_content_type("jpg");
    assert!(ct.contains("jpg"));
    assert!(ct.contains("image/jpeg"));
}

#[test]
fn test_generate_image_content_type_gif() {
    let ct = generate_image_content_type("gif");
    assert!(ct.contains("gif"));
    assert!(ct.contains("image/gif"));
}

#[test]
fn test_generate_image_content_type_bmp() {
    let ct = generate_image_content_type("bmp");
    assert!(ct.contains("bmp"));
    assert!(ct.contains("image/bmp"));
}

#[test]
fn test_generate_image_content_type_tiff() {
    let ct = generate_image_content_type("tiff");
    assert!(ct.contains("tiff"));
    assert!(ct.contains("image/tiff"));
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

#[test]
fn test_generate_image_xml_escape_special_chars() {
    let img = Image::new("photo & <test>.png", 1920000, 1080000, "PNG");
    let xml = generate_image_xml(&img, 1, 1);

    assert!(xml.contains("&amp;"));
    assert!(xml.contains("&lt;"));
    assert!(xml.contains("&gt;"));
}

#[test]
fn test_generate_image_xml_lock_aspect_ratio() {
    let img = Image::new("photo.png", 1920000, 1080000, "PNG");
    let xml = generate_image_xml(&img, 1, 1);

    assert!(xml.contains("noChangeAspect=\"1\""));
}

#[test]
fn test_generate_image_xml_fill_rect() {
    let img = Image::new("photo.png", 1920000, 1080000, "PNG");
    let xml = generate_image_xml(&img, 1, 1);

    assert!(xml.contains("a:fillRect"));
}

// ============================================================================
// VARIOUS IMAGE FORMATS
// ============================================================================

#[test]
fn test_image_various_formats() {
    let formats = vec!["PNG", "JPG", "JPEG", "GIF", "BMP", "TIFF", "SVG"];
    
    for format in formats {
        let img = Image::new("photo.ext", 1920, 1080, format);
        assert_eq!(img.format, format);
    }
}

#[test]
fn test_image_various_sizes() {
    let sizes = vec![
        (640, 480),
        (800, 600),
        (1024, 768),
        (1280, 720),
        (1920, 1080),
        (2560, 1440),
        (3840, 2160),
    ];
    
    for (width, height) in sizes {
        let img = Image::new("photo.png", width, height, "PNG");
        assert_eq!(img.width, width);
        assert_eq!(img.height, height);
    }
}

#[test]
fn test_image_square_formats() {
    let sizes = vec![512, 1024, 2048];
    
    for size in sizes {
        let img = Image::new("photo.png", size, size, "PNG");
        assert_eq!(img.aspect_ratio(), 1.0);
    }
}
