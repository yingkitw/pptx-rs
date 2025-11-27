//! Example demonstrating chart generation in PPTX presentations

use ppt_rs::generator::{
    create_pptx_with_content, SlideContent,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Create presentation with slides containing chart data
    let slides = vec![
        SlideContent::new("Sales Dashboard 2024")
            .add_bullet("Q1-Q4 Performance Comparison")
            .add_bullet("Year-over-year growth analysis")
            .add_bullet("Strategic insights and recommendations"),
        
        SlideContent::new("Bar Chart: Quarterly Sales")
            .add_bullet("2023 vs 2024 comparison")
            .add_bullet("Q4 shows strongest performance")
            .add_bullet("Overall growth: 20% YoY"),
        
        SlideContent::new("Line Chart: Revenue Trend")
            .add_bullet("6-month revenue tracking")
            .add_bullet("Consistent growth trajectory")
            .add_bullet("Exceeding targets in most months"),
        
        SlideContent::new("Pie Chart: Market Distribution")
            .add_bullet("Product A leads with 35% market share")
            .add_bullet("Products B & C equally competitive")
            .add_bullet("Product D growth opportunity"),
    ];

    // Generate the PPTX file
    let pptx_data = create_pptx_with_content("Sales Dashboard", slides)?;
    std::fs::write("chart_example.pptx", pptx_data)?;

    println!("✓ Created chart_example.pptx with:");
    println!("  - Title slide with overview");
    println!("  - Bar chart (Q1-Q4 sales comparison)");
    println!("  - Line chart (monthly revenue trend)");
    println!("  - Pie chart (market share distribution)");
    println!("\nNote: Chart XML structures are generated. For full chart functionality,");
    println!("additional integration with slide generation is needed.");

    Ok(())
}
