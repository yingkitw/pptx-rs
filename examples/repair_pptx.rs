//! Example demonstrating PPTX repair functionality
//!
//! This example shows how to:
//! 1. Open a potentially damaged PPTX file
//! 2. Validate and detect issues
//! 3. Repair the issues
//! 4. Save the repaired file

use ppt_rs::oxml::repair::{PptxRepair, RepairIssue};
use std::fs;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("PPTX Repair Example\n");
    println!("==================\n");

    // First, create a sample PPTX file to repair
    use ppt_rs::generator::{create_pptx_with_content, SlideContent};

    let slides = vec![
        SlideContent::new("Sample Presentation")
            .add_bullet("This is a sample slide")
            .add_bullet("Created for repair demonstration"),
        SlideContent::new("Second Slide")
            .add_bullet("More content here"),
    ];

    let pptx_data = create_pptx_with_content("Sample", slides)?;
    fs::write("sample_to_repair.pptx", &pptx_data)?;
    println!("Created sample PPTX file: sample_to_repair.pptx\n");

    // Open the file for repair
    let mut repair = PptxRepair::open("sample_to_repair.pptx")?;
    println!("Opened PPTX file for repair\n");

    // Validate and find issues
    println!("--- Validation ---");
    let issues = repair.validate();
    
    if issues.is_empty() {
        println!("No issues found! File is valid.\n");
    } else {
        println!("Found {} issue(s):\n", issues.len());
        for (i, issue) in issues.iter().enumerate() {
            println!("  {}. [Severity {}] {}", 
                i + 1, 
                issue.severity(),
                issue.description()
            );
            println!("     Repairable: {}", if issue.is_repairable() { "Yes" } else { "No" });
        }
        println!();
    }

    // Perform repair
    println!("--- Repair ---");
    let result = repair.repair();
    
    println!("Issues found: {}", result.total_issues());
    println!("Critical issues: {}", result.critical_issues());
    println!("Issues repaired: {}", result.issues_repaired.len());
    println!("Issues unrepaired: {}", result.issues_unrepaired.len());
    println!("File is now valid: {}", result.is_valid);
    println!();

    // Save the repaired file
    repair.save("repaired.pptx")?;
    println!("Saved repaired file: repaired.pptx\n");

    // Verify the repaired file
    println!("--- Verification ---");
    let mut verify = PptxRepair::open("repaired.pptx")?;
    let verify_issues = verify.validate();
    
    if verify_issues.is_empty() {
        println!("Repaired file is valid!");
    } else {
        println!("Repaired file still has {} issue(s)", verify_issues.len());
    }

    // Demonstrate issue types
    println!("\n--- Issue Types Reference ---");
    demonstrate_issue_types();

    // Cleanup
    fs::remove_file("sample_to_repair.pptx").ok();
    fs::remove_file("repaired.pptx").ok();

    println!("\nRepair example completed successfully!");
    Ok(())
}

fn demonstrate_issue_types() {
    let issue_examples = vec![
        RepairIssue::MissingPart {
            path: "[Content_Types].xml".to_string(),
            description: "Content types definition".to_string(),
        },
        RepairIssue::InvalidXml {
            path: "ppt/slides/slide1.xml".to_string(),
            error: "Malformed XML".to_string(),
        },
        RepairIssue::BrokenRelationship {
            source: "ppt/_rels/presentation.xml.rels".to_string(),
            target: "slides/slide99.xml".to_string(),
            rel_id: "rId99".to_string(),
        },
        RepairIssue::MissingSlideReference {
            slide_path: "ppt/slides/slide2.xml".to_string(),
        },
        RepairIssue::OrphanSlide {
            slide_path: "ppt/slides/slide99.xml".to_string(),
        },
    ];

    for issue in issue_examples {
        println!("  - {} (Severity: {}, Repairable: {})",
            match &issue {
                RepairIssue::MissingPart { .. } => "MissingPart",
                RepairIssue::InvalidXml { .. } => "InvalidXml",
                RepairIssue::BrokenRelationship { .. } => "BrokenRelationship",
                RepairIssue::MissingSlideReference { .. } => "MissingSlideReference",
                RepairIssue::OrphanSlide { .. } => "OrphanSlide",
                RepairIssue::InvalidContentType { .. } => "InvalidContentType",
                RepairIssue::CorruptedEntry { .. } => "CorruptedEntry",
                RepairIssue::MissingNamespace { .. } => "MissingNamespace",
                RepairIssue::EmptyRequiredElement { .. } => "EmptyRequiredElement",
            },
            issue.severity(),
            if issue.is_repairable() { "Yes" } else { "No" }
        );
    }
}
