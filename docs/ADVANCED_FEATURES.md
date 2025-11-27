# Advanced PPTX Features

This guide covers the advanced features available in PPTX-RS for creating rich presentations.

## Text Formatting

### Overview

The `TextFormat` struct provides comprehensive text styling options:
- **Bold** - Make text bold
- **Italic** - Make text italic
- **Underline** - Underline text
- **Color** - Set text color (RGB hex format)
- **Font Size** - Set font size in points

### Usage

```rust
use ppt_rs::generator::{TextFormat, FormattedText};

// Create a text format
let format = TextFormat::new()
    .bold()
    .italic()
    .color("FF0000")  // Red
    .font_size(24);

// Create formatted text
let text = FormattedText::new("Hello World")
    .bold()
    .color("0000FF")  // Blue
    .font_size(18);
```

### Color Format

Colors are specified as RGB hex values without the '#' prefix:
- `"FF0000"` - Red
- `"00FF00"` - Green
- `"0000FF"` - Blue
- `"FFFFFF"` - White
- `"000000"` - Black

### Font Sizes

Font sizes are specified in points (pt):
- `12` - Small text
- `18` - Normal text
- `24` - Large text
- `32` - Extra large text

## Shapes

### Overview

The shapes module provides support for creating various geometric shapes with customizable properties:
- **Rectangle** - Basic rectangular shape
- **Circle** - Round elliptical shape
- **Triangle** - Three-sided polygon
- **Diamond** - Rotated square
- **Arrow** - Directional indicator
- **Star** - Five-pointed star
- **Hexagon** - Six-sided polygon

### Shape Properties

Each shape can be customized with:
- **Position** - X and Y coordinates (in EMU)
- **Size** - Width and height (in EMU)
- **Fill** - Color and transparency
- **Line** - Border color and width
- **Text** - Text content inside the shape

### Usage

```rust
use ppt_rs::generator::{Shape, ShapeType, ShapeFill, ShapeLine};

// Create a rectangle with blue fill
let shape = Shape::new(ShapeType::Rectangle, 0, 0, 1000000, 500000)
    .with_fill(ShapeFill::new("0000FF"))
    .with_line(ShapeLine::new("000000", 25400))
    .with_text("Click Me");

// Create a circle with transparency
let circle = Shape::new(ShapeType::Circle, 1000000, 1000000, 500000, 500000)
    .with_fill(ShapeFill::new("FF0000").transparency(50));
```

### EMU (English Metric Units)

PPTX uses EMU for measurements:
- 1 inch = 914,400 EMU
- 1 cm = 360,000 EMU

Helper functions:
```rust
use ppt_rs::generator::shapes::{inches_to_emu, cm_to_emu, emu_to_inches};

let emu = inches_to_emu(1.0);      // 914400
let emu = cm_to_emu(2.54);         // 914400
let inches = emu_to_inches(914400); // 1.0
```

## Tables

### Overview

The tables module provides support for creating structured data tables with:
- **Rows** - Horizontal data containers
- **Cells** - Individual data cells
- **Formatting** - Cell colors and text styling
- **Layout** - Column widths and positioning

### Table Components

#### TableCell
```rust
use ppt_rs::generator::TableCell;

let cell = TableCell::new("Header")
    .bold()
    .background_color("0000FF");
```

#### TableRow
```rust
use ppt_rs::generator::{TableRow, TableCell};

let cells = vec![
    TableCell::new("Name"),
    TableCell::new("Age"),
];
let row = TableRow::new(cells).with_height(500000);
```

#### Table
```rust
use ppt_rs::generator::Table;

let table = Table::from_data(
    vec![
        vec!["Name", "Age"],
        vec!["Alice", "30"],
        vec!["Bob", "25"],
    ],
    vec![1000000, 1000000],  // Column widths
    0,                        // X position
    0,                        // Y position
);
```

### TableBuilder

For fluent API construction:

```rust
use ppt_rs::generator::TableBuilder;

let table = TableBuilder::new(vec![1000000, 1000000])
    .position(100000, 200000)
    .add_simple_row(vec!["Header 1", "Header 2"])
    .add_simple_row(vec!["Cell 1", "Cell 2"])
    .build();
```

## Examples

### Text Formatting Example

```rust
use ppt_rs::generator::{SlideContent, create_pptx_with_content};

let slides = vec![
    SlideContent::new("Formatted Text")
        .add_bullet("Bold text example")
        .add_bullet("Italic text example")
        .add_bullet("Colored text example"),
];

let pptx = create_pptx_with_content("Formatting Demo", slides)?;
```

### Shapes Example

```rust
use ppt_rs::generator::{Shape, ShapeType, ShapeFill};

let shapes = vec![
    Shape::new(ShapeType::Rectangle, 0, 0, 1000000, 500000)
        .with_fill(ShapeFill::new("FF0000")),
    Shape::new(ShapeType::Circle, 1500000, 0, 500000, 500000)
        .with_fill(ShapeFill::new("00FF00")),
];
```

### Tables Example

```rust
use ppt_rs::generator::TableBuilder;

let table = TableBuilder::new(vec![1000000, 1000000, 1000000])
    .add_simple_row(vec!["Q1", "Q2", "Q3"])
    .add_simple_row(vec!["$2.5M", "$3.1M", "$3.8M"])
    .build();
```

## Running the Advanced Features Example

```bash
cargo run --example advanced_features
```

This generates three presentations demonstrating:
1. Text formatting capabilities
2. Shape creation and styling
3. Table creation and formatting

## Integration with Slides

These features are designed to be integrated into the slide generation system. Future versions will support:
- Adding formatted text to slides
- Embedding shapes in presentations
- Creating data-driven tables
- Complex layouts with multiple elements

## Performance Considerations

- Text formatting is applied during XML generation
- Shapes are rendered as XML elements
- Tables generate complex XML structures
- Large tables may increase file size

## Troubleshooting

### Colors Not Appearing
- Ensure color format is correct (6-digit hex without #)
- Verify color is not the same as background

### Shapes Not Visible
- Check position and size values (in EMU)
- Ensure fill or line is specified
- Verify coordinates are within slide bounds

### Table Layout Issues
- Ensure column widths match number of columns
- Check cell text doesn't overflow
- Verify row heights are sufficient

## See Also

- [README.md](../README.md) - Main documentation
- [CODEBASE.md](../CODEBASE.md) - Code organization
- [examples/advanced_features.rs](../examples/advanced_features.rs) - Example code
