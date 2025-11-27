//! Example demonstrating shape creation in PPTX
//!
//! Shows various shape types, fills, lines, and text in shapes.

use pptx_rs::generator::{
    Shape, ShapeType, ShapeFill, ShapeLine,
    generate_shape_xml, generate_shapes_xml, generate_connector_xml,
    inches_to_emu, cm_to_emu,
};

fn main() {
    println!("╔════════════════════════════════════════════════════════════╗");
    println!("║         PPTX Shapes Demo                                   ║");
    println!("╚════════════════════════════════════════════════════════════╝\n");

    // =========================================================================
    // Basic Shapes
    // =========================================================================
    println!("📐 Basic Shapes:");
    
    let basic_shapes = [
        ShapeType::Rectangle,
        ShapeType::RoundedRectangle,
        ShapeType::Ellipse,
        ShapeType::Triangle,
        ShapeType::Diamond,
        ShapeType::Pentagon,
        ShapeType::Hexagon,
        ShapeType::Octagon,
    ];
    
    for shape_type in &basic_shapes {
        println!("   {} → {}", shape_type.display_name(), shape_type.preset_name());
    }

    // =========================================================================
    // Arrow Shapes
    // =========================================================================
    println!("\n➡️  Arrow Shapes:");
    
    let arrow_shapes = [
        ShapeType::RightArrow,
        ShapeType::LeftArrow,
        ShapeType::UpArrow,
        ShapeType::DownArrow,
        ShapeType::LeftRightArrow,
        ShapeType::UpDownArrow,
    ];
    
    for shape_type in &arrow_shapes {
        println!("   {} → {}", shape_type.display_name(), shape_type.preset_name());
    }

    // =========================================================================
    // Star and Banner Shapes
    // =========================================================================
    println!("\n⭐ Stars and Banners:");
    
    let star_shapes = [
        ShapeType::Star4,
        ShapeType::Star5,
        ShapeType::Star6,
        ShapeType::Star8,
        ShapeType::Ribbon,
        ShapeType::Wave,
    ];
    
    for shape_type in &star_shapes {
        println!("   {} → {}", shape_type.display_name(), shape_type.preset_name());
    }

    // =========================================================================
    // Callout Shapes
    // =========================================================================
    println!("\n💬 Callout Shapes:");
    
    let callout_shapes = [
        ShapeType::WedgeRectCallout,
        ShapeType::WedgeEllipseCallout,
        ShapeType::CloudCallout,
    ];
    
    for shape_type in &callout_shapes {
        println!("   {} → {}", shape_type.display_name(), shape_type.preset_name());
    }

    // =========================================================================
    // Flow Chart Shapes
    // =========================================================================
    println!("\n📊 Flow Chart Shapes:");
    
    let flowchart_shapes = [
        ShapeType::FlowChartProcess,
        ShapeType::FlowChartDecision,
        ShapeType::FlowChartTerminator,
        ShapeType::FlowChartDocument,
    ];
    
    for shape_type in &flowchart_shapes {
        println!("   {} → {}", shape_type.display_name(), shape_type.preset_name());
    }

    // =========================================================================
    // Other Shapes
    // =========================================================================
    println!("\n🎨 Other Shapes:");
    
    let other_shapes = [
        ShapeType::Heart,
        ShapeType::Lightning,
        ShapeType::Sun,
        ShapeType::Moon,
        ShapeType::Cloud,
    ];
    
    for shape_type in &other_shapes {
        println!("   {} → {}", shape_type.display_name(), shape_type.preset_name());
    }

    // =========================================================================
    // Shape with Fill
    // =========================================================================
    println!("\n🎨 Shape with Fill:");
    
    let filled_shape = Shape::new(
        ShapeType::Rectangle,
        inches_to_emu(1.0),
        inches_to_emu(1.0),
        inches_to_emu(3.0),
        inches_to_emu(2.0),
    ).with_fill(ShapeFill::new("4472C4")); // Blue fill
    
    let xml = generate_shape_xml(&filled_shape, 1);
    println!("   Generated XML ({} chars)", xml.len());
    println!("   Contains fill: {}", xml.contains("solidFill"));

    // =========================================================================
    // Shape with Line
    // =========================================================================
    println!("\n📏 Shape with Line:");
    
    let outlined_shape = Shape::new(
        ShapeType::Ellipse,
        inches_to_emu(1.0),
        inches_to_emu(1.0),
        inches_to_emu(2.0),
        inches_to_emu(2.0),
    ).with_line(ShapeLine::new("FF0000", 25400)); // Red outline, 2pt
    
    let xml = generate_shape_xml(&outlined_shape, 2);
    println!("   Generated XML ({} chars)", xml.len());
    println!("   Contains line: {}", xml.contains("a:ln"));

    // =========================================================================
    // Shape with Text
    // =========================================================================
    println!("\n📝 Shape with Text:");
    
    let text_shape = Shape::new(
        ShapeType::RoundedRectangle,
        cm_to_emu(5.0),
        cm_to_emu(3.0),
        cm_to_emu(8.0),
        cm_to_emu(4.0),
    )
    .with_fill(ShapeFill::new("70AD47")) // Green fill
    .with_text("Click Here!");
    
    let xml = generate_shape_xml(&text_shape, 3);
    println!("   Generated XML ({} chars)", xml.len());
    println!("   Contains text: {}", xml.contains("Click Here!"));

    // =========================================================================
    // Multiple Shapes
    // =========================================================================
    println!("\n📦 Multiple Shapes:");
    
    let shapes = vec![
        Shape::new(ShapeType::Rectangle, 0, 0, 1000000, 500000)
            .with_fill(ShapeFill::new("FF0000")),
        Shape::new(ShapeType::Ellipse, 1200000, 0, 500000, 500000)
            .with_fill(ShapeFill::new("00FF00")),
        Shape::new(ShapeType::Triangle, 1900000, 0, 500000, 500000)
            .with_fill(ShapeFill::new("0000FF")),
    ];
    
    let xml = generate_shapes_xml(&shapes, 10);
    println!("   Generated {} shapes", shapes.len());
    println!("   Total XML: {} chars", xml.len());

    // =========================================================================
    // Connector (Arrow Line)
    // =========================================================================
    println!("\n🔗 Connector:");
    
    let connector_xml = generate_connector_xml(
        0, 0,
        inches_to_emu(3.0), inches_to_emu(2.0),
        100,
        "000000",
        12700, // 1pt line
    );
    println!("   Generated connector XML ({} chars)", connector_xml.len());
    println!("   Has arrow head: {}", connector_xml.contains("triangle"));

    // =========================================================================
    // Summary
    // =========================================================================
    println!("\n╔════════════════════════════════════════════════════════════╗");
    println!("║                    Demo Complete                           ║");
    println!("╠════════════════════════════════════════════════════════════╣");
    println!("║  Shape Types Available: 40+                                ║");
    println!("║  Features:                                                 ║");
    println!("║  ✓ Basic shapes (rect, ellipse, triangle, etc.)            ║");
    println!("║  ✓ Arrow shapes (8 directions)                             ║");
    println!("║  ✓ Stars and banners                                       ║");
    println!("║  ✓ Callouts                                                ║");
    println!("║  ✓ Flow chart shapes                                       ║");
    println!("║  ✓ Fill colors with transparency                           ║");
    println!("║  ✓ Line/border styling                                     ║");
    println!("║  ✓ Text inside shapes                                      ║");
    println!("║  ✓ Connectors with arrow heads                             ║");
    println!("╚════════════════════════════════════════════════════════════╝");
}
