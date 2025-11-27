//! Example: Complete text styling with italic, underline, and colors
//!
//! Demonstrates all text formatting capabilities
//! Run with: cargo run --example text_styling_complete

use pptx_rs::generator::{SlideContent, create_pptx_with_content};
use std::fs;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔════════════════════════════════════════════════════════════╗");
    println!("║     Generating Complete Text Styling Examples             ║");
    println!("╚════════════════════════════════════════════════════════════╝");
    println!();

    fs::create_dir_all("examples/output")?;

    // Example 1: Italic and Underline
    println!("1. Creating presentation with italic and underline...");
    create_italic_underline_example()?;
    println!("   ✓ Created: examples/output/italic_underline.pptx");
    println!();

    // Example 2: Text Colors
    println!("2. Creating presentation with colored text...");
    create_colored_text_example()?;
    println!("   ✓ Created: examples/output/colored_text.pptx");
    println!();

    // Example 3: Combined Styling
    println!("3. Creating presentation with combined styling...");
    create_combined_styling_example()?;
    println!("   ✓ Created: examples/output/combined_styling.pptx");
    println!();

    // Example 4: Professional Presentation
    println!("4. Creating professional presentation...");
    create_professional_example()?;
    println!("   ✓ Created: examples/output/professional.pptx");
    println!();

    println!("✅ All text styling examples generated successfully!");
    println!();
    println!("Generated files:");
    println!("  - examples/output/italic_underline.pptx");
    println!("  - examples/output/colored_text.pptx");
    println!("  - examples/output/combined_styling.pptx");
    println!("  - examples/output/professional.pptx");
    println!();

    Ok(())
}

fn create_italic_underline_example() -> Result<(), Box<dyn std::error::Error>> {
    let slides = vec![
        SlideContent::new("Italic Text")
            .title_italic(true)
            .add_bullet("This is italic content")
            .add_bullet("More italic text here"),
        SlideContent::new("Underlined Text")
            .title_underline(true)
            .content_underline(true)
            .add_bullet("Underlined bullet point 1")
            .add_bullet("Underlined bullet point 2"),
        SlideContent::new("Combined Effects")
            .title_italic(true)
            .title_underline(true)
            .content_italic(true)
            .add_bullet("Italic and underlined content")
            .add_bullet("Multiple effects combined"),
    ];

    let pptx_data = create_pptx_with_content("Italic and Underline", slides)?;
    fs::write("examples/output/italic_underline.pptx", pptx_data)?;
    Ok(())
}

fn create_colored_text_example() -> Result<(), Box<dyn std::error::Error>> {
    let slides = vec![
        SlideContent::new("Red Title")
            .title_color("FF0000")
            .add_bullet("Red title text")
            .add_bullet("Regular content"),
        SlideContent::new("Blue Content")
            .content_color("0000FF")
            .add_bullet("Blue bullet point 1")
            .add_bullet("Blue bullet point 2")
            .add_bullet("Blue bullet point 3"),
        SlideContent::new("Green Title & Content")
            .title_color("00AA00")
            .content_color("00AA00")
            .add_bullet("Green title and content")
            .add_bullet("All text is green"),
        SlideContent::new("Purple Accent")
            .title_color("9933FF")
            .add_bullet("Purple title")
            .add_bullet("Regular content"),
    ];

    let pptx_data = create_pptx_with_content("Colored Text", slides)?;
    fs::write("examples/output/colored_text.pptx", pptx_data)?;
    Ok(())
}

fn create_combined_styling_example() -> Result<(), Box<dyn std::error::Error>> {
    let slides = vec![
        SlideContent::new("Bold & Italic & Red")
            .title_bold(true)
            .title_italic(true)
            .title_color("FF0000")
            .title_size(52)
            .add_bullet("Regular content"),
        SlideContent::new("Underlined Blue Content")
            .content_underline(true)
            .content_color("0000FF")
            .content_bold(true)
            .add_bullet("Bold, underlined, blue bullet")
            .add_bullet("Multiple effects applied"),
        SlideContent::new("Mixed Effects")
            .title_italic(true)
            .title_color("FF6600")
            .content_bold(true)
            .content_underline(true)
            .content_color("0066FF")
            .add_bullet("Title: italic, orange")
            .add_bullet("Content: bold, underlined, blue"),
        SlideContent::new("Professional Look")
            .title_bold(true)
            .title_size(48)
            .title_color("003366")
            .content_size(24)
            .add_bullet("Clean, professional styling")
            .add_bullet("Dark blue title with bold")
            .add_bullet("Larger content text"),
    ];

    let pptx_data = create_pptx_with_content("Combined Styling", slides)?;
    fs::write("examples/output/combined_styling.pptx", pptx_data)?;
    Ok(())
}

fn create_professional_example() -> Result<(), Box<dyn std::error::Error>> {
    let slides = vec![
        SlideContent::new("Company Presentation")
            .title_bold(true)
            .title_size(56)
            .title_color("003366")
            .content_size(32)
            .add_bullet("2025 Annual Review"),
        SlideContent::new("Key Highlights")
            .title_bold(true)
            .title_color("003366")
            .content_bold(true)
            .content_color("0066CC")
            .add_bullet("Revenue growth: +25%")
            .add_bullet("Market expansion: 3 new regions")
            .add_bullet("Team growth: +50 employees"),
        SlideContent::new("Strategic Initiatives")
            .title_italic(true)
            .title_color("FF6600")
            .content_underline(true)
            .add_bullet("Digital transformation")
            .add_bullet("Customer experience improvement")
            .add_bullet("Sustainability focus"),
        SlideContent::new("Q1 2025 Goals")
            .title_bold(true)
            .title_underline(true)
            .title_color("003366")
            .content_bold(true)
            .add_bullet("Launch new product line")
            .add_bullet("Expand to 5 new markets")
            .add_bullet("Achieve 30% revenue growth"),
        SlideContent::new("Thank You")
            .title_bold(true)
            .title_size(60)
            .title_color("003366")
            .content_italic(true)
            .content_size(28)
            .add_bullet("Questions & Discussion"),
    ];

    let pptx_data = create_pptx_with_content("Professional Presentation", slides)?;
    fs::write("examples/output/professional.pptx", pptx_data)?;
    Ok(())
}
