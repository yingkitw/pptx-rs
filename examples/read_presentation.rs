//! Example demonstrating PPTX reading and content extraction
//!
//! This example shows how to:
//! - Open and read existing PPTX files
//! - Extract presentation metadata
//! - Parse slide content (titles, bullets, shapes)
//! - Extract all text from a presentation

use ppt_rs::generator::{create_pptx_with_content, SlideContent, SlideLayout};
use ppt_rs::oxml::{PresentationReader, SlideParser};
use std::fs;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔════════════════════════════════════════════════════════════╗");
    println!("║         PPTX Reading & Parsing Demo                        ║");
    println!("╚════════════════════════════════════════════════════════════╝\n");

    // =========================================================================
    // Step 1: Create a sample PPTX to read
    // =========================================================================
    println!("📝 Step 1: Creating sample presentation...");
    
    let slides = vec![
        SlideContent::new("Welcome to PPTX-RS")
            .layout(SlideLayout::CenteredTitle)
            .title_bold(true)
            .title_color("1F497D"),
        
        SlideContent::new("Features Overview")
            .add_bullet("Create presentations programmatically")
            .add_bullet("Read existing PPTX files")
            .add_bullet("Extract text and metadata")
            .add_bullet("Parse shapes and tables"),
        
        SlideContent::new("Technical Details")
            .layout(SlideLayout::TwoColumn)
            .add_bullet("XML parsing with xml-rs")
            .add_bullet("ZIP handling with zip crate")
            .add_bullet("ECMA-376 compliant")
            .add_bullet("Rust 2024 edition")
            .add_bullet("Cross-platform")
            .add_bullet("No external dependencies"),
        
        SlideContent::new("Summary")
            .add_bullet("Full read/write support")
            .add_bullet("Comprehensive API")
            .add_bullet("Well tested"),
    ];
    
    let pptx_data = create_pptx_with_content("PPTX-RS Demo", slides)?;
    fs::write("sample_presentation.pptx", &pptx_data)?;
    println!("   ✓ Created sample_presentation.pptx ({} bytes)\n", pptx_data.len());

    // =========================================================================
    // Step 2: Open and read the presentation
    // =========================================================================
    println!("📖 Step 2: Opening presentation...");
    
    let reader = PresentationReader::open("sample_presentation.pptx")?;
    let info = reader.info();
    
    println!("   Presentation Info:");
    println!("   ├── Title: {}", info.title.as_deref().unwrap_or("(none)"));
    println!("   ├── Creator: {}", info.creator.as_deref().unwrap_or("(none)"));
    println!("   ├── Slides: {}", info.slide_count);
    println!("   └── Revision: {}\n", info.revision.unwrap_or(0));

    // =========================================================================
    // Step 3: Parse each slide
    // =========================================================================
    println!("📑 Step 3: Parsing slides...");
    
    for i in 0..reader.slide_count() {
        let slide = reader.get_slide(i)?;
        
        println!("\n   Slide {}:", i + 1);
        println!("   ├── Title: {}", slide.title.as_deref().unwrap_or("(none)"));
        println!("   ├── Shapes: {}", slide.shapes.len());
        println!("   ├── Tables: {}", slide.tables.len());
        
        if !slide.body_text.is_empty() {
            println!("   └── Body text:");
            for (j, text) in slide.body_text.iter().enumerate() {
                let prefix = if j == slide.body_text.len() - 1 { "       └──" } else { "       ├──" };
                println!("{}  {}", prefix, text);
            }
        } else {
            println!("   └── Body text: (none)");
        }
    }

    // =========================================================================
    // Step 4: Extract all text
    // =========================================================================
    println!("\n📋 Step 4: Extracting all text...");
    
    let all_text = reader.extract_all_text()?;
    println!("   Found {} text items:", all_text.len());
    for (i, text) in all_text.iter().take(10).enumerate() {
        println!("   {}. {}", i + 1, text);
    }
    if all_text.len() > 10 {
        println!("   ... and {} more", all_text.len() - 10);
    }

    // =========================================================================
    // Step 5: Direct XML parsing (advanced)
    // =========================================================================
    println!("\n🔧 Step 5: Direct XML parsing (advanced)...");
    
    // You can also parse slide XML directly
    let sample_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
           xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cSld>
            <p:spTree>
                <p:sp>
                    <p:nvSpPr>
                        <p:cNvPr id="2" name="Title"/>
                        <p:nvPr><p:ph type="title"/></p:nvPr>
                    </p:nvSpPr>
                    <p:txBody>
                        <a:p>
                            <a:r>
                                <a:rPr b="1" sz="4400"/>
                                <a:t>Direct Parse Example</a:t>
                            </a:r>
                        </a:p>
                    </p:txBody>
                </p:sp>
            </p:spTree>
        </p:cSld>
    </p:sld>"#;
    
    let parsed = SlideParser::parse(sample_xml)?;
    println!("   Parsed XML directly:");
    println!("   ├── Title: {}", parsed.title.as_deref().unwrap_or("(none)"));
    println!("   └── Shapes: {}", parsed.shapes.len());
    
    if let Some(shape) = parsed.shapes.first() {
        if let Some(para) = shape.paragraphs.first() {
            if let Some(run) = para.runs.first() {
                println!("\n   Text formatting detected:");
                println!("   ├── Bold: {}", run.bold);
                println!("   ├── Font size: {:?}", run.font_size);
                println!("   └── Text: {}", run.text);
            }
        }
    }

    // Cleanup
    fs::remove_file("sample_presentation.pptx").ok();

    // =========================================================================
    // Summary
    // =========================================================================
    println!("\n╔════════════════════════════════════════════════════════════╗");
    println!("║                    Demo Complete                           ║");
    println!("╠════════════════════════════════════════════════════════════╣");
    println!("║  Capabilities Demonstrated:                                ║");
    println!("║  ✓ PresentationReader::open() - Open PPTX files            ║");
    println!("║  ✓ reader.info() - Get presentation metadata               ║");
    println!("║  ✓ reader.get_slide(i) - Parse individual slides           ║");
    println!("║  ✓ reader.extract_all_text() - Extract all text            ║");
    println!("║  ✓ SlideParser::parse() - Direct XML parsing               ║");
    println!("╚════════════════════════════════════════════════════════════╝");

    Ok(())
}
