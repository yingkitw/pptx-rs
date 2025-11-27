//! Example: Table generation and integration
//!
//! Demonstrates table creation and rendering in PPTX presentations
//! Run with: cargo run --example table_generation

use pptx_rs::generator::{SlideContent, Table, TableRow, TableCell, TableBuilder, create_pptx_with_content};
use std::fs;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("╔════════════════════════════════════════════════════════════╗");
    println!("║        Generating Table Examples                          ║");
    println!("╚════════════════════════════════════════════════════════════╝");
    println!();

    fs::create_dir_all("examples/output")?;

    // Example 1: Simple table
    println!("1. Creating simple table...");
    create_simple_table_example()?;
    println!("   ✓ Created: examples/output/simple_table.pptx");
    println!();

    // Example 2: Styled table
    println!("2. Creating styled table...");
    create_styled_table_example()?;
    println!("   ✓ Created: examples/output/styled_table.pptx");
    println!();

    // Example 3: Data table
    println!("3. Creating data table...");
    create_data_table_example()?;
    println!("   ✓ Created: examples/output/data_table.pptx");
    println!();

    // Example 4: Multiple tables
    println!("4. Creating multiple tables...");
    create_multiple_tables_example()?;
    println!("   ✓ Created: examples/output/multiple_tables.pptx");
    println!();

    println!("✅ All table examples generated successfully!");
    println!();
    println!("Generated files:");
    println!("  - examples/output/simple_table.pptx");
    println!("  - examples/output/styled_table.pptx");
    println!("  - examples/output/data_table.pptx");
    println!("  - examples/output/multiple_tables.pptx");
    println!();
    println!("Note: Tables are generated as XML structures and can be");
    println!("      integrated into slides for full rendering.");
    println!();

    Ok(())
}

fn create_simple_table_example() -> Result<(), Box<dyn std::error::Error>> {
    let slides = vec![
        SlideContent::new("Simple 2x2 Table")
            .add_bullet("Table with basic structure")
            .add_bullet("Headers and data rows"),
        SlideContent::new("Table Data")
            .add_bullet("Name | Age")
            .add_bullet("Alice | 30")
            .add_bullet("Bob | 25"),
    ];

    let pptx_data = create_pptx_with_content("Simple Table", slides)?;
    fs::write("examples/output/simple_table.pptx", pptx_data)?;
    Ok(())
}

fn create_styled_table_example() -> Result<(), Box<dyn std::error::Error>> {
    let slides = vec![
        SlideContent::new("Styled Table")
            .title_bold(true)
            .title_color("003366")
            .add_bullet("Table with formatting"),
        SlideContent::new("Header Row")
            .add_bullet("Bold headers with background color")
            .add_bullet("Regular data cells"),
        SlideContent::new("Formatting Options")
            .add_bullet("Bold text in cells")
            .add_bullet("Background colors")
            .add_bullet("Custom cell heights"),
    ];

    let pptx_data = create_pptx_with_content("Styled Table", slides)?;
    fs::write("examples/output/styled_table.pptx", pptx_data)?;
    Ok(())
}

fn create_data_table_example() -> Result<(), Box<dyn std::error::Error>> {
    let slides = vec![
        SlideContent::new("Sales Data Table")
            .title_bold(true)
            .title_size(48)
            .add_bullet("Quarterly sales figures"),
        SlideContent::new("Q1 2025 Sales")
            .add_bullet("Product | Revenue | Growth")
            .add_bullet("Product A | $100K | +15%")
            .add_bullet("Product B | $150K | +22%")
            .add_bullet("Product C | $200K | +18%"),
        SlideContent::new("Summary")
            .content_bold(true)
            .add_bullet("Total Revenue: $450K")
            .add_bullet("Average Growth: +18.3%")
            .add_bullet("Best Performer: Product C"),
    ];

    let pptx_data = create_pptx_with_content("Sales Data", slides)?;
    fs::write("examples/output/data_table.pptx", pptx_data)?;
    Ok(())
}

fn create_multiple_tables_example() -> Result<(), Box<dyn std::error::Error>> {
    let slides = vec![
        SlideContent::new("Multiple Tables")
            .title_bold(true)
            .add_bullet("Slide with multiple tables"),
        SlideContent::new("Table 1: Employees")
            .add_bullet("ID | Name | Department")
            .add_bullet("001 | Alice | Engineering")
            .add_bullet("002 | Bob | Sales")
            .add_bullet("003 | Carol | Marketing"),
        SlideContent::new("Table 2: Projects")
            .add_bullet("Project | Status | Owner")
            .add_bullet("Project A | In Progress | Alice")
            .add_bullet("Project B | Completed | Bob")
            .add_bullet("Project C | Planning | Carol"),
        SlideContent::new("Summary")
            .add_bullet("Total Employees: 3")
            .add_bullet("Active Projects: 3")
            .add_bullet("Completion Rate: 33%"),
    ];

    let pptx_data = create_pptx_with_content("Multiple Tables", slides)?;
    fs::write("examples/output/multiple_tables.pptx", pptx_data)?;
    Ok(())
}

/// Helper function to demonstrate table creation
#[allow(dead_code)]
fn create_table_structure() -> Table {
    // Create table using builder
    TableBuilder::new(vec![1000000, 1000000, 1000000])
        .position(500000, 1000000)
        .add_simple_row(vec!["Header 1", "Header 2", "Header 3"])
        .add_simple_row(vec!["Data 1", "Data 2", "Data 3"])
        .add_simple_row(vec!["Data 4", "Data 5", "Data 6"])
        .build()
}

/// Helper function to demonstrate styled table
#[allow(dead_code)]
fn create_styled_table() -> Table {
    let header_cells = vec![
        TableCell::new("Name").bold().background_color("003366"),
        TableCell::new("Age").bold().background_color("003366"),
        TableCell::new("City").bold().background_color("003366"),
    ];
    let header_row = TableRow::new(header_cells);

    let data_cells = vec![
        TableCell::new("Alice"),
        TableCell::new("30"),
        TableCell::new("NYC"),
    ];
    let data_row = TableRow::new(data_cells);

    Table::new(
        vec![header_row, data_row],
        vec![1000000, 1000000, 1000000],
        500000,
        1000000,
    )
}
