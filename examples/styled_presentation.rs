//! Example: Styled presentation with custom fonts and formatting
//!
//! Demonstrates text formatting capabilities
//! Run with: cargo run --example styled_presentation

use ppt_rs::generator::{SlideContent, create_pptx_with_content};
use std::fs;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔════════════════════════════════════════════════════════════╗");
    println!("║     Generating Styled Presentation Examples               ║");
    println!("╚════════════════════════════════════════════════════════════╝");
    println!();

    // Create output directory
    fs::create_dir_all("examples/output")?;

    // Example 1: Large title with small content
    println!("1. Creating presentation with custom font sizes...");
    create_large_title_example()?;
    println!("   ✓ Created: examples/output/large_title.pptx");
    println!();

    // Example 2: Bold content
    println!("2. Creating presentation with bold formatting...");
    create_bold_content_example()?;
    println!("   ✓ Created: examples/output/bold_content.pptx");
    println!();

    // Example 3: Mixed formatting
    println!("3. Creating presentation with mixed formatting...");
    create_mixed_formatting_example()?;
    println!("   ✓ Created: examples/output/mixed_formatting.pptx");
    println!();

    println!("✅ All styled presentations generated successfully!");
    println!();
    println!("Generated files:");
    println!("  - examples/output/large_title.pptx");
    println!("  - examples/output/bold_content.pptx");
    println!("  - examples/output/mixed_formatting.pptx");
    println!();
    println!("Open these files to see the text formatting in action!");
    println!();

    Ok(())
}

fn create_large_title_example() -> Result<(), Box<dyn std::error::Error>> {
    let slides = vec![
        SlideContent::new("Large Title Slide")
            .title_size(60)  // 60pt title
            .content_size(16) // 16pt content
            .add_bullet("Small bullet point 1")
            .add_bullet("Small bullet point 2")
            .add_bullet("Small bullet point 3"),
        SlideContent::new("Regular Formatting")
            .add_bullet("Default 44pt title")
            .add_bullet("Default 28pt content")
            .add_bullet("Standard formatting"),
    ];

    let pptx_data = create_pptx_with_content("Large Title Example", slides)?;
    fs::write("examples/output/large_title.pptx", pptx_data)?;
    Ok(())
}

fn create_bold_content_example() -> Result<(), Box<dyn std::error::Error>> {
    let slides = vec![
        SlideContent::new("Bold Content Slide")
            .title_bold(true)
            .content_bold(true)  // Make bullets bold
            .add_bullet("Bold bullet point 1")
            .add_bullet("Bold bullet point 2")
            .add_bullet("Bold bullet point 3"),
        SlideContent::new("Regular Content")
            .title_bold(true)
            .content_bold(false)  // Regular bullets
            .add_bullet("Regular bullet point 1")
            .add_bullet("Regular bullet point 2"),
    ];

    let pptx_data = create_pptx_with_content("Bold Content Example", slides)?;
    fs::write("examples/output/bold_content.pptx", pptx_data)?;
    Ok(())
}

fn create_mixed_formatting_example() -> Result<(), Box<dyn std::error::Error>> {
    let slides = vec![
        SlideContent::new("Title Slide")
            .title_size(52)
            .title_bold(true)
            .content_size(24)
            .content_bold(false)
            .add_bullet("Large content text")
            .add_bullet("Still readable")
            .add_bullet("Professional look"),
        SlideContent::new("Compact Slide")
            .title_size(36)
            .title_bold(false)
            .content_size(18)
            .content_bold(true)
            .add_bullet("Smaller title")
            .add_bullet("Bold content")
            .add_bullet("Tight spacing"),
        SlideContent::new("Summary")
            .title_size(48)
            .title_bold(true)
            .content_size(32)
            .content_bold(true)
            .add_bullet("Large bold text")
            .add_bullet("High impact")
            .add_bullet("Great for emphasis"),
    ];

    let pptx_data = create_pptx_with_content("Mixed Formatting Example", slides)?;
    fs::write("examples/output/mixed_formatting.pptx", pptx_data)?;
    Ok(())
}
