//! Example demonstrating PPTX editing capabilities
//!
//! This example shows how to:
//! - Open existing PPTX files for editing
//! - Add new slides to a presentation
//! - Update existing slide content
//! - Remove slides
//! - Save modified presentations

use ppt_rs::generator::{create_pptx_with_content, SlideContent, SlideLayout};
use ppt_rs::oxml::{PresentationEditor, PresentationReader};
use std::fs;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔════════════════════════════════════════════════════════════╗");
    println!("║         PPTX Editing Demo                                  ║");
    println!("╚════════════════════════════════════════════════════════════╝\n");

    // =========================================================================
    // Step 1: Create an initial presentation
    // =========================================================================
    println!("📝 Step 1: Creating initial presentation...");
    
    let initial_slides = vec![
        SlideContent::new("Original Presentation")
            .layout(SlideLayout::CenteredTitle)
            .title_bold(true)
            .title_color("1F497D"),
        
        SlideContent::new("Slide 1: Introduction")
            .add_bullet("This is the original content")
            .add_bullet("Created programmatically"),
        
        SlideContent::new("Slide 2: Features")
            .add_bullet("Feature A")
            .add_bullet("Feature B")
            .add_bullet("Feature C"),
    ];
    
    let pptx_data = create_pptx_with_content("Original Presentation", initial_slides)?;
    fs::write("original.pptx", &pptx_data)?;
    println!("   ✓ Created original.pptx with 3 slides\n");

    // =========================================================================
    // Step 2: Open and inspect the presentation
    // =========================================================================
    println!("📖 Step 2: Opening presentation for editing...");
    
    let mut editor = PresentationEditor::open("original.pptx")?;
    println!("   ✓ Opened original.pptx");
    println!("   ├── Slide count: {}", editor.slide_count());
    
    // Read first slide
    let slide0 = editor.get_slide(0)?;
    println!("   └── First slide title: {:?}\n", slide0.title);

    // =========================================================================
    // Step 3: Add new slides
    // =========================================================================
    println!("➕ Step 3: Adding new slides...");
    
    // Add a new slide at the end
    let new_slide1 = SlideContent::new("New Slide: Added via Editor")
        .add_bullet("This slide was added programmatically")
        .add_bullet("Using PresentationEditor")
        .add_bullet("After the presentation was created")
        .title_color("9BBB59");
    
    let idx1 = editor.add_slide(new_slide1)?;
    println!("   ✓ Added slide at index {}", idx1);
    
    // Add another slide
    let new_slide2 = SlideContent::new("Another New Slide")
        .layout(SlideLayout::TwoColumn)
        .add_bullet("Left column item 1")
        .add_bullet("Left column item 2")
        .add_bullet("Right column item 1")
        .add_bullet("Right column item 2");
    
    let idx2 = editor.add_slide(new_slide2)?;
    println!("   ✓ Added slide at index {}", idx2);
    println!("   └── Total slides now: {}\n", editor.slide_count());

    // =========================================================================
    // Step 4: Update existing slide
    // =========================================================================
    println!("✏️  Step 4: Updating existing slide...");
    
    let updated_slide = SlideContent::new("Slide 2: Updated Features")
        .add_bullet("Feature A - Enhanced!")
        .add_bullet("Feature B - Improved!")
        .add_bullet("Feature C - Optimized!")
        .add_bullet("Feature D - NEW!")
        .title_color("C0504D")
        .content_bold(true);
    
    editor.update_slide(2, updated_slide)?;
    println!("   ✓ Updated slide at index 2\n");

    // =========================================================================
    // Step 5: Save modified presentation
    // =========================================================================
    println!("💾 Step 5: Saving modified presentation...");
    
    editor.save("modified.pptx")?;
    println!("   ✓ Saved as modified.pptx\n");

    // =========================================================================
    // Step 6: Verify the changes
    // =========================================================================
    println!("🔍 Step 6: Verifying changes...");
    
    let reader = PresentationReader::open("modified.pptx")?;
    println!("   Modified presentation:");
    println!("   ├── Slide count: {}", reader.slide_count());
    
    for i in 0..reader.slide_count() {
        let slide = reader.get_slide(i)?;
        let title = slide.title.as_deref().unwrap_or("(no title)");
        let bullets = slide.body_text.len();
        println!("   {}── Slide {}: \"{}\" ({} bullets)", 
                 if i == reader.slide_count() - 1 { "└" } else { "├" },
                 i + 1, 
                 title,
                 bullets);
    }

    // =========================================================================
    // Step 7: Demonstrate slide removal
    // =========================================================================
    println!("\n🗑️  Step 7: Demonstrating slide removal...");
    
    let mut editor2 = PresentationEditor::open("modified.pptx")?;
    println!("   Before removal: {} slides", editor2.slide_count());
    
    // Remove the last slide
    editor2.remove_slide(editor2.slide_count() - 1)?;
    println!("   ✓ Removed last slide");
    println!("   After removal: {} slides", editor2.slide_count());
    
    editor2.save("trimmed.pptx")?;
    println!("   ✓ Saved as trimmed.pptx");

    // Cleanup
    fs::remove_file("original.pptx").ok();
    fs::remove_file("modified.pptx").ok();
    fs::remove_file("trimmed.pptx").ok();

    // =========================================================================
    // Summary
    // =========================================================================
    println!("\n╔════════════════════════════════════════════════════════════╗");
    println!("║                    Demo Complete                           ║");
    println!("╠════════════════════════════════════════════════════════════╣");
    println!("║  Capabilities Demonstrated:                                ║");
    println!("║  ✓ PresentationEditor::open() - Open for editing           ║");
    println!("║  ✓ editor.add_slide() - Add new slides                     ║");
    println!("║  ✓ editor.update_slide() - Modify existing slides          ║");
    println!("║  ✓ editor.remove_slide() - Remove slides                   ║");
    println!("║  ✓ editor.save() - Save modified presentation              ║");
    println!("║  ✓ editor.get_slide() - Read slide content                 ║");
    println!("╚════════════════════════════════════════════════════════════╝");

    Ok(())
}
