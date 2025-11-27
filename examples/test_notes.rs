//! Test notes functionality with proper GUID

use ppt_rs::generator::{create_pptx_with_content, SlideContent, SlideLayout};
use std::fs;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating test presentation with speaker notes...");
    
    let slides = vec![
        SlideContent::new("Slide 1 - With Notes")
            .layout(SlideLayout::CenteredTitle)
            .title_size(54)
            .title_bold(true)
            .notes("This is a speaker note for slide 1. It should appear in presenter view."),
        
        SlideContent::new("Slide 2 - Also With Notes")
            .add_bullet("Bullet point 1")
            .add_bullet("Bullet point 2")
            .add_bullet("Bullet point 3")
            .notes("Notes for slide 2 with bullet points."),
        
        SlideContent::new("Slide 3 - No Notes")
            .add_bullet("This slide has no notes"),
        
        SlideContent::new("Slide 4 - More Notes")
            .layout(SlideLayout::TwoColumn)
            .add_bullet("Left 1")
            .add_bullet("Left 2")
            .add_bullet("Right 1")
            .add_bullet("Right 2")
            .notes("Two column layout with speaker notes."),
    ];
    
    let pptx_data = create_pptx_with_content("Test Notes", slides)?;
    fs::write("test_notes.pptx", &pptx_data)?;
    println!("Created test_notes.pptx ({} bytes)", pptx_data.len());
    println!("\nOpen in PowerPoint and check presenter view for notes!");
    
    Ok(())
}
