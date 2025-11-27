//! Example: Image handling and XML generation
//!
//! Demonstrates image creation, scaling, and XML generation
//! Run with: cargo run --example image_handling

use pptx_rs::generator::{Image, ImageBuilder, generate_image_xml};
use std::fs;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔════════════════════════════════════════════════════════════╗");
    println!("║        Generating Image Handling Examples                 ║");
    println!("╚════════════════════════════════════════════════════════════╝");
    println!();

    fs::create_dir_all("examples/output")?;

    // Example 1: Image creation
    println!("1. Creating image metadata...");
    demonstrate_image_creation();
    println!("   ✓ Image metadata created");
    println!();

    // Example 2: Image scaling
    println!("2. Demonstrating image scaling...");
    demonstrate_image_scaling();
    println!("   ✓ Image scaling demonstrated");
    println!();

    // Example 3: Image builder
    println!("3. Using image builder...");
    demonstrate_image_builder();
    println!("   ✓ Image builder demonstrated");
    println!();

    // Example 4: Image XML generation
    println!("4. Generating image XML...");
    demonstrate_image_xml_generation();
    println!("   ✓ Image XML generated");
    println!();

    // Example 5: Multiple image formats
    println!("5. Supporting multiple formats...");
    demonstrate_multiple_formats();
    println!("   ✓ Multiple formats supported");
    println!();

    println!("✅ All image handling examples completed!");
    println!();
    println!("Features demonstrated:");
    println!("  - Image metadata creation");
    println!("  - Image scaling and aspect ratios");
    println!("  - Builder pattern for images");
    println!("  - XML generation for images");
    println!("  - Multiple image format support");
    println!();

    Ok(())
}

fn demonstrate_image_creation() {
    let img = Image::new("photo.png", 1920, 1080, "PNG");
    println!("   Image: {}", img.filename);
    println!("   Size: {}x{}", img.width, img.height);
    println!("   Format: {}", img.format);
    println!("   Aspect Ratio: {:.2}", img.aspect_ratio());
}

fn demonstrate_image_scaling() {
    let original = Image::new("photo.png", 1920, 1080, "PNG");
    println!("   Original: {}x{}", original.width, original.height);

    let scaled_width = original.clone().scale_to_width(960);
    println!("   Scaled to width 960: {}x{}", scaled_width.width, scaled_width.height);

    let scaled_height = original.clone().scale_to_height(540);
    println!("   Scaled to height 540: {}x{}", scaled_height.width, scaled_height.height);
}

fn demonstrate_image_builder() {
    let img = ImageBuilder::new("photo.png", 1920, 1080)
        .position(500000, 1000000)
        .scale_to_width(960)
        .build();

    println!("   Filename: {}", img.filename);
    println!("   Size: {}x{}", img.width, img.height);
    println!("   Position: ({}, {})", img.x, img.y);
    println!("   MIME Type: {}", img.mime_type());
}

fn demonstrate_image_xml_generation() {
    let img = Image::new("photo.png", 1920000, 1080000, "PNG")
        .position(500000, 1000000);

    let xml = generate_image_xml(&img, 1, 1);
    println!("   Generated XML length: {} bytes", xml.len());
    println!("   Contains p:pic: {}", xml.contains("p:pic"));
    println!("   Contains a:blip: {}", xml.contains("a:blip"));
    println!("   Contains position: {}", xml.contains("x=\"500000\""));
}

fn demonstrate_multiple_formats() {
    let formats = vec![
        ("photo.png", "PNG"),
        ("photo.jpg", "JPG"),
        ("photo.gif", "GIF"),
        ("photo.bmp", "BMP"),
        ("photo.tiff", "TIFF"),
    ];

    for (filename, format) in formats {
        let img = Image::new(filename, 1920, 1080, format);
        println!("   {}: {}", filename, img.mime_type());
    }
}
