//! Example: Table generation and integration
//!
//! Demonstrates table creation and rendering in PPTX presentations
//! Run with: cargo run --example table_generation

use ppt_rs::generator::{SlideContent, Table, TableRow, TableCell, TableBuilder, create_pptx_with_content};
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
    let table = Table::from_data(
        vec![
            vec!["Name", "Age"],
            vec!["Alice", "30"],
            vec!["Bob", "25"],
        ],
        vec![2000000, 2000000],
        500000,
        1500000,
    );

    let slides = vec![
        SlideContent::new("Simple 2x2 Table")
            .add_bullet("Table with basic structure")
            .add_bullet("Headers and data rows"),
        SlideContent::new("Table Data")
            .table(table),
    ];

    let pptx_data = create_pptx_with_content("Simple Table", slides)?;
    fs::write("examples/output/simple_table.pptx", pptx_data)?;
    Ok(())
}

fn create_styled_table_example() -> Result<(), Box<dyn std::error::Error>> {
    let header_cells = vec![
        TableCell::new("Name").bold().background_color("003366"),
        TableCell::new("Age").bold().background_color("003366"),
        TableCell::new("City").bold().background_color("003366"),
    ];
    let header_row = TableRow::new(header_cells);

    let data_rows = vec![
        TableRow::new(vec![
            TableCell::new("Alice"),
            TableCell::new("30"),
            TableCell::new("NYC"),
        ]),
        TableRow::new(vec![
            TableCell::new("Bob"),
            TableCell::new("28"),
            TableCell::new("LA"),
        ]),
        TableRow::new(vec![
            TableCell::new("Carol"),
            TableCell::new("35"),
            TableCell::new("Chicago"),
        ]),
    ];

    let table = Table::new(
        vec![vec![header_row], data_rows].concat(),
        vec![1500000, 1500000, 1500000],
        500000,
        1500000,
    );

    let slides = vec![
        SlideContent::new("Styled Table")
            .title_bold(true)
            .title_color("003366")
            .add_bullet("Table with formatting"),
        SlideContent::new("People Data")
            .table(table),
    ];

    let pptx_data = create_pptx_with_content("Styled Table", slides)?;
    fs::write("examples/output/styled_table.pptx", pptx_data)?;
    Ok(())
}

fn create_data_table_example() -> Result<(), Box<dyn std::error::Error>> {
    let header_cells = vec![
        TableCell::new("Product").bold().background_color("1F497D"),
        TableCell::new("Revenue").bold().background_color("1F497D"),
        TableCell::new("Growth").bold().background_color("1F497D"),
    ];
    let header_row = TableRow::new(header_cells);

    let data_rows = vec![
        TableRow::new(vec![
            TableCell::new("Product A"),
            TableCell::new("$100K"),
            TableCell::new("+15%"),
        ]),
        TableRow::new(vec![
            TableCell::new("Product B"),
            TableCell::new("$150K"),
            TableCell::new("+22%"),
        ]),
        TableRow::new(vec![
            TableCell::new("Product C"),
            TableCell::new("$200K"),
            TableCell::new("+18%"),
        ]),
    ];

    let table = Table::new(
        vec![vec![header_row], data_rows].concat(),
        vec![2000000, 2000000, 1500000],
        457200,
        1400000,
    );

    let slides = vec![
        SlideContent::new("Sales Data Table")
            .title_bold(true)
            .title_size(48)
            .add_bullet("Quarterly sales figures"),
        SlideContent::new("Q1 2025 Sales")
            .table(table),
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
    // Table 1: Employees
    let emp_header = TableRow::new(vec![
        TableCell::new("ID").bold().background_color("4F81BD"),
        TableCell::new("Name").bold().background_color("4F81BD"),
        TableCell::new("Department").bold().background_color("4F81BD"),
    ]);
    let emp_rows = vec![
        TableRow::new(vec![
            TableCell::new("001"),
            TableCell::new("Alice"),
            TableCell::new("Engineering"),
        ]),
        TableRow::new(vec![
            TableCell::new("002"),
            TableCell::new("Bob"),
            TableCell::new("Sales"),
        ]),
        TableRow::new(vec![
            TableCell::new("003"),
            TableCell::new("Carol"),
            TableCell::new("Marketing"),
        ]),
    ];
    let emp_table = Table::new(
        vec![vec![emp_header], emp_rows].concat(),
        vec![1000000, 2000000, 2000000],
        500000,
        1500000,
    );

    // Table 2: Projects
    let proj_header = TableRow::new(vec![
        TableCell::new("Project").bold().background_color("003366"),
        TableCell::new("Status").bold().background_color("003366"),
        TableCell::new("Owner").bold().background_color("003366"),
    ]);
    let proj_rows = vec![
        TableRow::new(vec![
            TableCell::new("Project A"),
            TableCell::new("In Progress"),
            TableCell::new("Alice"),
        ]),
        TableRow::new(vec![
            TableCell::new("Project B"),
            TableCell::new("Completed"),
            TableCell::new("Bob"),
        ]),
        TableRow::new(vec![
            TableCell::new("Project C"),
            TableCell::new("Planning"),
            TableCell::new("Carol"),
        ]),
    ];
    let proj_table = Table::new(
        vec![vec![proj_header], proj_rows].concat(),
        vec![2000000, 2000000, 1500000],
        500000,
        1500000,
    );

    let slides = vec![
        SlideContent::new("Multiple Tables")
            .title_bold(true)
            .add_bullet("Slide with multiple tables"),
        SlideContent::new("Table 1: Employees")
            .table(emp_table),
        SlideContent::new("Table 2: Projects")
            .table(proj_table),
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
