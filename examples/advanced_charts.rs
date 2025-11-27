//! Advanced chart examples with multiple chart types and data

use ppt_rs::generator::{
    create_pptx_with_content, SlideContent,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Slide 1: Title
    let title_slide = SlideContent::new("Advanced Analytics Dashboard")
        .add_bullet("Comprehensive data visualization")
        .add_bullet("Multiple chart types and formats")
        .add_bullet("Interactive insights and trends");

    // Slide 2: Bar Chart - Sales by Region
    let bar_slide = SlideContent::new("Regional Sales Performance")
        .add_bullet("East region leads with consistent growth")
        .add_bullet("Q3 shows strongest performance across all regions")
        .add_bullet("North region demonstrates 28% growth Q1-Q3");

    // Slide 3: Line Chart - Customer Growth
    let line_slide = SlideContent::new("Customer Acquisition & Retention")
        .add_bullet("200% growth in new customer acquisition")
        .add_bullet("Retention rate improved from 85% to 93.5%")
        .add_bullet("Accelerating growth in Q4");

    // Slide 4: Pie Chart - Product Mix
    let pie_slide = SlideContent::new("Product Mix Analysis")
        .add_bullet("Enterprise segment: 42% of revenue")
        .add_bullet("Mid-Market: 28% with strong growth potential")
        .add_bullet("SMB & Startup: 30% combined, high growth rate");

    // Slide 5: Multiple metrics
    let metrics_slide = SlideContent::new("Key Performance Indicators")
        .add_bullet("Total Revenue: $2.8M (↑ 35% YoY)")
        .add_bullet("Customer Count: 450 (↑ 200% YoY)")
        .add_bullet("Retention Rate: 93.5% (↑ 8.5 points)")
        .add_bullet("Average Deal Size: $6,222 (↑ 12% YoY)")
        .add_bullet("Sales Cycle: 45 days (↓ 15% faster)");

    // Slide 6: Forecast
    let forecast_slide = SlideContent::new("2025 Forecast & Goals")
        .add_bullet("Projected Revenue: $4.2M (50% growth)")
        .add_bullet("Target Customer Count: 750")
        .add_bullet("Goal Retention Rate: 95%")
        .add_bullet("Expansion into 3 new markets")
        .add_bullet("Launch 2 new product lines");

    let slides = vec![
        title_slide,
        bar_slide,
        line_slide,
        pie_slide,
        metrics_slide,
        forecast_slide,
    ];

    // Generate the PPTX file
    let pptx_data = create_pptx_with_content("Analytics Dashboard", slides)?;
    std::fs::write("advanced_charts.pptx", pptx_data)?;

    println!("✓ Created advanced_charts.pptx with:");
    println!("  - Title slide");
    println!("  - Bar chart: Regional sales performance");
    println!("  - Line chart: Customer growth trajectory");
    println!("  - Pie chart: Revenue by product category");
    println!("  - KPI summary slide");
    println!("  - Forecast and goals slide");
    println!("\nTotal slides: 6");

    Ok(())
}
