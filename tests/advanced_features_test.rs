//! Integration tests for advanced PPTX features
//!
//! Tests text formatting, shapes, and tables functionality
//! Also generates example PPTX files for manual verification

use pptx_rs::generator::{
    TextFormat, FormattedText,
    Shape, ShapeType, ShapeFill, ShapeLine,
    Table, TableRow, TableCell, TableBuilder,
    SlideContent, create_pptx_with_content,
};
use std::fs;
use std::path::Path;

// ============================================================================
// TEXT FORMATTING TESTS
// ============================================================================

#[test]
fn test_text_format_default() {
    let format = TextFormat::new();
    assert!(!format.bold);
    assert!(!format.italic);
    assert!(!format.underline);
    assert!(format.color.is_none());
    assert!(format.font_size.is_none());
}

#[test]
fn test_text_format_bold() {
    let format = TextFormat::new().bold();
    assert!(format.bold);
    assert!(!format.italic);
}

#[test]
fn test_text_format_italic() {
    let format = TextFormat::new().italic();
    assert!(format.italic);
    assert!(!format.bold);
}

#[test]
fn test_text_format_underline() {
    let format = TextFormat::new().underline();
    assert!(format.underline);
}

#[test]
fn test_text_format_color() {
    let format = TextFormat::new().color("FF0000");
    assert_eq!(format.color, Some("FF0000".to_string()));
}

#[test]
fn test_text_format_color_lowercase() {
    let format = TextFormat::new().color("ff0000");
    assert_eq!(format.color, Some("FF0000".to_string()));
}

#[test]
fn test_text_format_font_size() {
    let format = TextFormat::new().font_size(24);
    assert_eq!(format.font_size, Some(24));
}

#[test]
fn test_text_format_combined() {
    let format = TextFormat::new()
        .bold()
        .italic()
        .underline()
        .color("0000FF")
        .font_size(18);

    assert!(format.bold);
    assert!(format.italic);
    assert!(format.underline);
    assert_eq!(format.color, Some("0000FF".to_string()));
    assert_eq!(format.font_size, Some(18));
}

#[test]
fn test_formatted_text_default() {
    let text = FormattedText::new("Hello");
    assert_eq!(text.text, "Hello");
    assert!(!text.format.bold);
}

#[test]
fn test_formatted_text_with_format() {
    let format = TextFormat::new().bold().color("FF0000");
    let text = FormattedText::new("World").with_format(format);
    assert_eq!(text.text, "World");
    assert!(text.format.bold);
    assert_eq!(text.format.color, Some("FF0000".to_string()));
}

#[test]
fn test_formatted_text_builder() {
    let text = FormattedText::new("Test")
        .bold()
        .italic()
        .color("00FF00")
        .font_size(32);

    assert_eq!(text.text, "Test");
    assert!(text.format.bold);
    assert!(text.format.italic);
    assert_eq!(text.format.color, Some("00FF00".to_string()));
    assert_eq!(text.format.font_size, Some(32));
}

// ============================================================================
// SHAPE TESTS
// ============================================================================

#[test]
fn test_shape_type_rectangle() {
    assert_eq!(ShapeType::Rectangle.preset_name(), "rect");
}

#[test]
fn test_shape_type_circle() {
    assert_eq!(ShapeType::Circle.preset_name(), "ellipse");
}

#[test]
fn test_shape_type_triangle() {
    assert_eq!(ShapeType::Triangle.preset_name(), "triangle");
}

#[test]
fn test_shape_type_diamond() {
    assert_eq!(ShapeType::Diamond.preset_name(), "diamond");
}

#[test]
fn test_shape_type_arrow() {
    assert_eq!(ShapeType::Arrow.preset_name(), "rightArrow");
}

#[test]
fn test_shape_type_star() {
    assert_eq!(ShapeType::Star.preset_name(), "star5");
}

#[test]
fn test_shape_type_hexagon() {
    assert_eq!(ShapeType::Hexagon.preset_name(), "hexagon");
}

#[test]
fn test_shape_fill_default() {
    let fill = ShapeFill::new("FF0000");
    assert_eq!(fill.color, "FF0000");
    assert!(fill.transparency.is_none());
}

#[test]
fn test_shape_fill_with_transparency() {
    let fill = ShapeFill::new("00FF00").transparency(50);
    assert_eq!(fill.color, "00FF00");
    assert_eq!(fill.transparency, Some(50000));
}

#[test]
fn test_shape_fill_full_transparency() {
    let fill = ShapeFill::new("0000FF").transparency(100);
    assert_eq!(fill.transparency, Some(0));
}

#[test]
fn test_shape_fill_no_transparency() {
    let fill = ShapeFill::new("FFFFFF").transparency(0);
    assert_eq!(fill.transparency, Some(100000));
}

#[test]
fn test_shape_line_default() {
    let line = ShapeLine::new("000000", 25400);
    assert_eq!(line.color, "000000");
    assert_eq!(line.width, 25400);
}

#[test]
fn test_shape_line_lowercase_color() {
    let line = ShapeLine::new("ff0000", 50000);
    assert_eq!(line.color, "FF0000");
    assert_eq!(line.width, 50000);
}

#[test]
fn test_shape_basic() {
    let shape = Shape::new(ShapeType::Rectangle, 0, 0, 1000000, 500000);
    assert_eq!(shape.x, 0);
    assert_eq!(shape.y, 0);
    assert_eq!(shape.width, 1000000);
    assert_eq!(shape.height, 500000);
    assert!(shape.fill.is_none());
    assert!(shape.line.is_none());
    assert!(shape.text.is_none());
}

#[test]
fn test_shape_with_fill() {
    let fill = ShapeFill::new("0000FF");
    let shape = Shape::new(ShapeType::Circle, 100000, 200000, 500000, 500000)
        .with_fill(fill);

    assert_eq!(shape.x, 100000);
    assert_eq!(shape.y, 200000);
    assert!(shape.fill.is_some());
    assert_eq!(shape.fill.unwrap().color, "0000FF");
}

#[test]
fn test_shape_with_line() {
    let line = ShapeLine::new("FF0000", 25400);
    let shape = Shape::new(ShapeType::Diamond, 0, 0, 800000, 800000)
        .with_line(line);

    assert!(shape.line.is_some());
    let line_ref = shape.line.as_ref().unwrap();
    assert_eq!(line_ref.color, "FF0000");
    assert_eq!(line_ref.width, 25400);
}

#[test]
fn test_shape_with_text() {
    let shape = Shape::new(ShapeType::Rectangle, 0, 0, 1000000, 500000)
        .with_text("Click Me");

    assert!(shape.text.is_some());
    assert_eq!(shape.text.unwrap(), "Click Me");
}

#[test]
fn test_shape_complete() {
    let shape = Shape::new(ShapeType::Arrow, 500000, 500000, 1000000, 500000)
        .with_fill(ShapeFill::new("FF0000"))
        .with_line(ShapeLine::new("000000", 25400))
        .with_text("Next");

    assert_eq!(shape.x, 500000);
    assert_eq!(shape.y, 500000);
    assert_eq!(shape.width, 1000000);
    assert_eq!(shape.height, 500000);
    assert!(shape.fill.is_some());
    assert!(shape.line.is_some());
    assert_eq!(shape.text, Some("Next".to_string()));
}

// ============================================================================
// TABLE TESTS
// ============================================================================

#[test]
fn test_table_cell_default() {
    let cell = TableCell::new("Data");
    assert_eq!(cell.text, "Data");
    assert!(!cell.bold);
    assert!(cell.background_color.is_none());
}

#[test]
fn test_table_cell_bold() {
    let cell = TableCell::new("Header").bold();
    assert!(cell.bold);
}

#[test]
fn test_table_cell_background_color() {
    let cell = TableCell::new("Cell").background_color("0000FF");
    assert_eq!(cell.background_color, Some("0000FF".to_string()));
}

#[test]
fn test_table_cell_combined() {
    let cell = TableCell::new("Header")
        .bold()
        .background_color("FF0000");

    assert_eq!(cell.text, "Header");
    assert!(cell.bold);
    assert_eq!(cell.background_color, Some("FF0000".to_string()));
}

#[test]
fn test_table_row_default() {
    let cells = vec![
        TableCell::new("A"),
        TableCell::new("B"),
    ];
    let row = TableRow::new(cells);

    assert_eq!(row.cells.len(), 2);
    assert!(row.height.is_none());
}

#[test]
fn test_table_row_with_height() {
    let cells = vec![TableCell::new("Data")];
    let row = TableRow::new(cells).with_height(500000);

    assert_eq!(row.height, Some(500000));
}

#[test]
fn test_table_basic() {
    let rows = vec![
        TableRow::new(vec![TableCell::new("A"), TableCell::new("B")]),
        TableRow::new(vec![TableCell::new("1"), TableCell::new("2")]),
    ];
    let table = Table::new(rows, vec![1000000, 1000000], 0, 0);

    assert_eq!(table.row_count(), 2);
    assert_eq!(table.column_count(), 2);
    assert_eq!(table.x, 0);
    assert_eq!(table.y, 0);
}

#[test]
fn test_table_from_data() {
    let data = vec![
        vec!["Name", "Age"],
        vec!["Alice", "30"],
        vec!["Bob", "25"],
    ];
    let table = Table::from_data(data, vec![1000000, 1000000], 100000, 200000);

    assert_eq!(table.row_count(), 3);
    assert_eq!(table.column_count(), 2);
    assert_eq!(table.x, 100000);
    assert_eq!(table.y, 200000);
}

#[test]
fn test_table_builder_empty() {
    let table = TableBuilder::new(vec![1000000, 1000000]).build();

    assert_eq!(table.row_count(), 0);
    assert_eq!(table.column_count(), 2);
}

#[test]
fn test_table_builder_with_rows() {
    let table = TableBuilder::new(vec![1000000, 1000000])
        .add_simple_row(vec!["Header 1", "Header 2"])
        .add_simple_row(vec!["Cell 1", "Cell 2"])
        .build();

    assert_eq!(table.row_count(), 2);
    assert_eq!(table.column_count(), 2);
}

#[test]
fn test_table_builder_position() {
    let table = TableBuilder::new(vec![1000000])
        .position(500000, 1000000)
        .build();

    assert_eq!(table.x, 500000);
    assert_eq!(table.y, 1000000);
}

#[test]
fn test_table_builder_complete() {
    let table = TableBuilder::new(vec![1000000, 1000000, 1000000])
        .position(100000, 200000)
        .add_simple_row(vec!["Q1", "Q2", "Q3"])
        .add_simple_row(vec!["$2.5M", "$3.1M", "$3.8M"])
        .add_simple_row(vec!["$4.2M", "$5.1M", "$6.0M"])
        .build();

    assert_eq!(table.row_count(), 3);
    assert_eq!(table.column_count(), 3);
    assert_eq!(table.x, 100000);
    assert_eq!(table.y, 200000);
}

// ============================================================================
// EMU CONVERSION TESTS
// ============================================================================

#[test]
fn test_inches_to_emu() {
    use pptx_rs::generator::shapes::inches_to_emu;
    assert_eq!(inches_to_emu(1.0), 914400);
    assert_eq!(inches_to_emu(2.0), 1828800);
    assert_eq!(inches_to_emu(0.5), 457200);
}

#[test]
fn test_emu_to_inches() {
    use pptx_rs::generator::shapes::emu_to_inches;
    let inches = emu_to_inches(914400);
    assert!((inches - 1.0).abs() < 0.001);
}

#[test]
fn test_cm_to_emu() {
    use pptx_rs::generator::shapes::cm_to_emu;
    assert_eq!(cm_to_emu(2.54), 914400); // 1 inch
    assert_eq!(cm_to_emu(5.08), 1828800); // 2 inches
}

#[test]
fn test_emu_conversions_roundtrip() {
    use pptx_rs::generator::shapes::{inches_to_emu, emu_to_inches};
    
    let original = 1.5;
    let emu = inches_to_emu(original);
    let result = emu_to_inches(emu);
    
    assert!((result - original).abs() < 0.001);
}

// ============================================================================
// EDGE CASES AND VALIDATION
// ============================================================================

#[test]
fn test_text_format_multiple_applications() {
    let format = TextFormat::new()
        .bold()
        .bold()  // Apply twice
        .italic()
        .italic(); // Apply twice

    assert!(format.bold);
    assert!(format.italic);
}

#[test]
fn test_shape_zero_dimensions() {
    let shape = Shape::new(ShapeType::Rectangle, 0, 0, 0, 0);
    assert_eq!(shape.width, 0);
    assert_eq!(shape.height, 0);
}

#[test]
fn test_shape_large_dimensions() {
    let shape = Shape::new(ShapeType::Rectangle, 0, 0, 9144000, 6858000);
    assert_eq!(shape.width, 9144000);
    assert_eq!(shape.height, 6858000);
}

#[test]
fn test_table_single_cell() {
    let table = Table::from_data(
        vec![vec!["Single"]],
        vec![1000000],
        0,
        0,
    );

    assert_eq!(table.row_count(), 1);
    assert_eq!(table.column_count(), 1);
}

#[test]
fn test_table_many_columns() {
    let table = Table::from_data(
        vec![vec!["A", "B", "C", "D", "E"]],
        vec![1000000; 5],
        0,
        0,
    );

    assert_eq!(table.row_count(), 1);
    assert_eq!(table.column_count(), 5);
}

#[test]
fn test_table_many_rows() {
    let data: Vec<Vec<&str>> = (0..10)
        .map(|i| vec![match i {
            0 => "Row 0",
            1 => "Row 1",
            2 => "Row 2",
            3 => "Row 3",
            4 => "Row 4",
            5 => "Row 5",
            6 => "Row 6",
            7 => "Row 7",
            8 => "Row 8",
            _ => "Row 9",
        }])
        .collect();
    
    let table = Table::from_data(data, vec![1000000], 0, 0);
    assert_eq!(table.row_count(), 10);
}

// ============================================================================
// PPTX GENERATION TESTS - Generate example files for manual verification
// ============================================================================

#[test]
fn test_generate_text_formatting_pptx() {
    // Create output directory
    let _ = fs::create_dir_all("target/test_output");

    let slides = vec![
        SlideContent::new("Text Formatting Demo")
            .title_size(52)
            .content_size(24)
            .add_bullet("Large 24pt content")
            .add_bullet("Easy to read")
            .add_bullet("Professional appearance"),
        SlideContent::new("Bold Content")
            .title_bold(true)
            .content_bold(true)
            .add_bullet("Bold bullet 1")
            .add_bullet("Bold bullet 2")
            .add_bullet("Bold bullet 3"),
        SlideContent::new("Mixed Formatting")
            .title_size(48)
            .title_bold(true)
            .content_size(20)
            .content_bold(false)
            .add_bullet("Large title, medium content")
            .add_bullet("Different font sizes")
            .add_bullet("Varied formatting"),
    ];

    let result = create_pptx_with_content("Text Formatting Test", slides);
    assert!(result.is_ok());

    let pptx_data = result.unwrap();
    let output_path = "target/test_output/test_text_formatting.pptx";
    let write_result = fs::write(output_path, pptx_data);
    assert!(write_result.is_ok());
    assert!(Path::new(output_path).exists());
}

#[test]
fn test_generate_simple_pptx() {
    let _ = fs::create_dir_all("target/test_output");

    let slides = vec![
        SlideContent::new("Introduction")
            .add_bullet("Welcome")
            .add_bullet("Agenda"),
        SlideContent::new("Key Points")
            .add_bullet("Point 1")
            .add_bullet("Point 2")
            .add_bullet("Point 3"),
        SlideContent::new("Conclusion")
            .add_bullet("Summary")
            .add_bullet("Thank you"),
    ];

    let result = create_pptx_with_content("Simple Test", slides);
    assert!(result.is_ok());

    let pptx_data = result.unwrap();
    let output_path = "target/test_output/test_simple.pptx";
    let write_result = fs::write(output_path, pptx_data);
    assert!(write_result.is_ok());
    assert!(Path::new(output_path).exists());
}

#[test]
fn test_generate_large_presentation() {
    let _ = fs::create_dir_all("target/test_output");

    let mut slides = Vec::new();
    for i in 1..=5 {
        let slide = SlideContent::new(&format!("Slide {}", i))
            .title_size(44 + (i as u32 * 2))
            .add_bullet(&format!("Content for slide {}", i))
            .add_bullet("Additional information")
            .add_bullet("More details");
        slides.push(slide);
    }

    let result = create_pptx_with_content("Large Presentation Test", slides);
    assert!(result.is_ok());

    let pptx_data = result.unwrap();
    let output_path = "target/test_output/test_large_presentation.pptx";
    let write_result = fs::write(output_path, pptx_data);
    assert!(write_result.is_ok());
    assert!(Path::new(output_path).exists());
}

#[test]
fn test_generate_styled_presentation() {
    let _ = fs::create_dir_all("target/test_output");

    let slides = vec![
        SlideContent::new("Large Title")
            .title_size(60)
            .content_size(16)
            .add_bullet("Small 16pt content")
            .add_bullet("Compact layout")
            .add_bullet("High density"),
        SlideContent::new("Regular Title")
            .title_size(44)
            .content_size(28)
            .add_bullet("Standard 28pt content")
            .add_bullet("Default formatting")
            .add_bullet("Balanced layout"),
        SlideContent::new("Bold Emphasis")
            .title_bold(true)
            .content_bold(true)
            .title_size(48)
            .content_size(26)
            .add_bullet("Bold title and content")
            .add_bullet("Strong emphasis")
            .add_bullet("High impact"),
    ];

    let result = create_pptx_with_content("Styled Presentation Test", slides);
    assert!(result.is_ok());

    let pptx_data = result.unwrap();
    let output_path = "target/test_output/test_styled_presentation.pptx";
    let write_result = fs::write(output_path, pptx_data);
    assert!(write_result.is_ok());
    assert!(Path::new(output_path).exists());
}

#[test]
fn test_generate_all_examples() {
    let _ = fs::create_dir_all("target/test_output");

    // Generate multiple example presentations
    let examples = vec![
        ("Text Formatting", vec![
            SlideContent::new("Formatting Demo")
                .title_size(50)
                .content_size(22)
                .add_bullet("Custom font sizes")
                .add_bullet("Professional styling"),
        ]),
        ("Bold Content", vec![
            SlideContent::new("Bold Slide")
                .title_bold(true)
                .content_bold(true)
                .add_bullet("Bold text")
                .add_bullet("Strong emphasis"),
        ]),
        ("Mixed Styles", vec![
            SlideContent::new("Slide 1")
                .title_size(48)
                .content_size(24)
                .add_bullet("First slide"),
            SlideContent::new("Slide 2")
                .title_size(44)
                .content_size(28)
                .add_bullet("Second slide"),
        ]),
    ];

    for (name, slides) in examples {
        let result = create_pptx_with_content(name, slides);
        assert!(result.is_ok());

        let pptx_data = result.unwrap();
        let filename = format!("target/test_output/test_{}.pptx", name.to_lowercase().replace(" ", "_"));
        let write_result = fs::write(&filename, pptx_data);
        assert!(write_result.is_ok());
        assert!(Path::new(&filename).exists());
    }
}
