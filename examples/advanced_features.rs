//! Example: Advanced PPTX features
//!
//! Demonstrates text formatting, shapes, and tables
//! Run with: cargo run --example advanced_features

use ppt_rs::generator::{
    SlideContent,
    create_pptx_with_content,
};
use std::fs;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔════════════════════════════════════════════════════════════╗");
    println!("║     Generating Advanced PPTX Examples                     ║");
    println!("╚════════════════════════════════════════════════════════════╝");
    println!();

    // Create output directory
    fs::create_dir_all("examples/output")?;

    // Example 1: Text Formatting
    println!("1. Creating presentation with text formatting...");
    create_text_formatting_example()?;
    println!("   ✓ Created: examples/output/text_formatting.pptx");
    println!();

    // Example 2: Shapes
    println!("2. Creating presentation with shapes...");
    create_shapes_example()?;
    println!("   ✓ Created: examples/output/shapes_demo.pptx");
    println!();

    // Example 3: Tables
    println!("3. Creating presentation with tables...");
    create_tables_example()?;
    println!("   ✓ Created: examples/output/tables_demo.pptx");
    println!();

    println!("✅ All advanced examples generated successfully!");
    println!();
    println!("Generated files:");
    println!("  - examples/output/text_formatting.pptx");
    println!("  - examples/output/shapes_demo.pptx");
    println!("  - examples/output/tables_demo.pptx");
    println!();

    Ok(())
}

fn create_text_formatting_example() -> Result<(), Box<dyn std::error::Error>> {
    let slides = vec![
        SlideContent::new("Text Formatting Examples")
            .add_bullet("Bold text demonstration")
            .add_bullet("Italic text demonstration")
            .add_bullet("Underlined text demonstration")
            .add_bullet("Colored text demonstration"),
        SlideContent::new("Font Sizes")
            .add_bullet("Small text (12pt)")
            .add_bullet("Normal text (18pt)")
            .add_bullet("Large text (24pt)")
            .add_bullet("Extra large text (32pt)"),
        SlideContent::new("Combined Formatting")
            .add_bullet("Bold and italic together")
            .add_bullet("Colored and underlined text")
            .add_bullet("Large bold red text")
            .add_bullet("Small italic blue text"),
    ];

    let pptx_data = create_pptx_with_content("Text Formatting", slides)?;
    fs::write("examples/output/text_formatting.pptx", pptx_data)?;
    Ok(())
}

fn create_shapes_example() -> Result<(), Box<dyn std::error::Error>> {
    let slides = vec![
        SlideContent::new("Shape Types")
            .add_bullet("Rectangle - basic rectangular shape")
            .add_bullet("Circle - round elliptical shape")
            .add_bullet("Triangle - three-sided polygon")
            .add_bullet("Diamond - rotated square")
            .add_bullet("Arrow - directional indicator")
            .add_bullet("Star - five-pointed star")
            .add_bullet("Hexagon - six-sided polygon"),
        SlideContent::new("Shape Properties")
            .add_bullet("Fill colors - solid color fills")
            .add_bullet("Transparency - adjustable opacity")
            .add_bullet("Borders - customizable lines")
            .add_bullet("Text - add text inside shapes")
            .add_bullet("Positioning - precise placement")
            .add_bullet("Sizing - flexible dimensions"),
        SlideContent::new("Shape Examples")
            .add_bullet("Rectangle: 2x1 inch blue box")
            .add_bullet("Circle: 1 inch red circle")
            .add_bullet("Triangle: green triangle with border")
            .add_bullet("Diamond: yellow diamond with text")
            .add_bullet("Arrow: orange arrow pointing right"),
    ];

    let pptx_data = create_pptx_with_content("Shape Demonstrations", slides)?;
    fs::write("examples/output/shapes_demo.pptx", pptx_data)?;
    Ok(())
}

fn create_tables_example() -> Result<(), Box<dyn std::error::Error>> {
    let slides = vec![
        SlideContent::new("Table Basics")
            .add_bullet("Tables organize data in rows and columns")
            .add_bullet("Headers can be formatted differently")
            .add_bullet("Cells can have background colors")
            .add_bullet("Text can be bold or styled")
            .add_bullet("Column widths are customizable"),
        SlideContent::new("Sales Data Example")
            .add_bullet("Q1 2025: $2.5M revenue")
            .add_bullet("Q2 2025: $3.1M revenue")
            .add_bullet("Q3 2025: $3.8M revenue")
            .add_bullet("Q4 2025: $4.2M revenue (projected)")
            .add_bullet("Total: $13.6M annual revenue"),
        SlideContent::new("Team Members")
            .add_bullet("Alice Johnson - Engineering Lead")
            .add_bullet("Bob Smith - Product Manager")
            .add_bullet("Carol Davis - Design Lead")
            .add_bullet("David Wilson - Marketing Manager")
            .add_bullet("Eve Martinez - Operations"),
    ];

    let pptx_data = create_pptx_with_content("Table Examples", slides)?;
    fs::write("examples/output/tables_demo.pptx", pptx_data)?;
    Ok(())
}
