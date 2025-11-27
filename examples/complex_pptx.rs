//! Example: Generate complex PPTX files with content
//!
//! Run with: cargo run --example complex_pptx

use std::fs;
use ppt_rs::generator::{SlideContent, create_pptx_with_content};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Generating complex PPTX files...\n");

    // Create output directory
    fs::create_dir_all("examples/output")?;

    // Example 1: Business presentation with multiple slides
    println!("Creating business_presentation.pptx...");
    let slides = vec![
        SlideContent::new("Business Overview")
            .add_bullet("Company mission and vision")
            .add_bullet("Market position")
            .add_bullet("Strategic goals"),
        SlideContent::new("Q1 2025 Results")
            .add_bullet("Revenue: $2.5M (+15% YoY)")
            .add_bullet("Customer acquisition: 250 new clients")
            .add_bullet("Market share increased to 12%"),
        SlideContent::new("Key Initiatives")
            .add_bullet("Product roadmap expansion")
            .add_bullet("Team growth and hiring")
            .add_bullet("International market entry"),
        SlideContent::new("Financial Projections")
            .add_bullet("Q2 target: $2.8M revenue")
            .add_bullet("Expected 20% growth")
            .add_bullet("Break-even by Q3"),
    ];

    let pptx_data = create_pptx_with_content("Business Presentation 2025", slides)?;
    fs::write("examples/output/business_presentation.pptx", pptx_data)?;
    println!("✓ Created: examples/output/business_presentation.pptx\n");

    // Example 2: Training presentation
    println!("Creating training_course.pptx...");
    let slides = vec![
        SlideContent::new("Rust Programming Basics")
            .add_bullet("Introduction to Rust")
            .add_bullet("Why Rust matters")
            .add_bullet("Course objectives"),
        SlideContent::new("Ownership & Borrowing")
            .add_bullet("Understanding ownership")
            .add_bullet("Borrowing and references")
            .add_bullet("Lifetimes explained"),
        SlideContent::new("Error Handling")
            .add_bullet("Result and Option types")
            .add_bullet("Error propagation")
            .add_bullet("Best practices"),
        SlideContent::new("Concurrency")
            .add_bullet("Threads and channels")
            .add_bullet("Async/await patterns")
            .add_bullet("Practical examples"),
        SlideContent::new("Conclusion")
            .add_bullet("Key takeaways")
            .add_bullet("Resources for learning")
            .add_bullet("Next steps"),
    ];

    let pptx_data = create_pptx_with_content("Rust Programming Course", slides)?;
    fs::write("examples/output/training_course.pptx", pptx_data)?;
    println!("✓ Created: examples/output/training_course.pptx\n");

    // Example 3: Project proposal
    println!("Creating project_proposal.pptx...");
    let slides = vec![
        SlideContent::new("Project Proposal: PPTX Generator")
            .add_bullet("Objective: Create a Rust library for PPTX generation")
            .add_bullet("Timeline: 3 months")
            .add_bullet("Budget: $50,000"),
        SlideContent::new("Problem Statement")
            .add_bullet("Existing Python library is slow")
            .add_bullet("Need for type-safe solution")
            .add_bullet("Performance requirements not met"),
        SlideContent::new("Proposed Solution")
            .add_bullet("Rust-based PPTX generator")
            .add_bullet("10x performance improvement")
            .add_bullet("Full API compatibility"),
        SlideContent::new("Implementation Plan")
            .add_bullet("Phase 1: Core ZIP handling (Month 1)")
            .add_bullet("Phase 2: XML generation (Month 2)")
            .add_bullet("Phase 3: Testing & optimization (Month 3)"),
        SlideContent::new("Expected Outcomes")
            .add_bullet("Production-ready library")
            .add_bullet("Comprehensive documentation")
            .add_bullet("100% test coverage"),
        SlideContent::new("Questions & Discussion"),
    ];

    let pptx_data = create_pptx_with_content("PPTX Generator Project", slides)?;
    fs::write("examples/output/project_proposal.pptx", pptx_data)?;
    println!("✓ Created: examples/output/project_proposal.pptx\n");

    println!("✅ All complex PPTX files generated successfully!");
    println!("\nGenerated files:");
    println!("  - examples/output/business_presentation.pptx (4 slides with content)");
    println!("  - examples/output/training_course.pptx (5 slides with content)");
    println!("  - examples/output/project_proposal.pptx (6 slides with content)");

    // Verify files
    println!("\nVerifying PPTX format:");
    for file in &[
        "examples/output/business_presentation.pptx",
        "examples/output/training_course.pptx",
        "examples/output/project_proposal.pptx",
    ] {
        let metadata = fs::metadata(file)?;
        let size = metadata.len();
        println!("  ✓ {} ({} bytes)", file, size);
    }

    Ok(())
}
