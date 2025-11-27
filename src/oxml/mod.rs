//! OXML (Office XML) element handling
//!
//! Provides XML parsing and content extraction for Office Open XML documents.

pub mod action;
pub mod chart;
pub mod coreprops;
pub mod dml;
pub mod editor;
pub mod ns;
pub mod presentation;
pub mod shapes;
pub mod simpletypes;
pub mod slide;
pub mod table;
pub mod text;
pub mod theme;
pub mod xmlchemy;

// Core XML parsing
pub use xmlchemy::{XmlElement, XmlParser, BaseOxmlElement};

// Slide parsing
pub use slide::{SlideParser, ParsedSlide, ParsedShape, ParsedTable, ParsedTableCell, Paragraph, TextRun};

// Presentation reading
pub use presentation::{PresentationReader, PresentationInfo};

// Presentation editing
pub use editor::PresentationEditor;

// Namespace utilities
pub use ns::Namespace;

// Text elements
pub use text::{TextBody, TextParagraph, TextRun as OxmlTextRun, RunProperties, ParagraphProperties, BodyProperties};

// Table elements
pub use table::{Table as OxmlTable, TableRow as OxmlTableRow, TableCell as OxmlTableCell, GridColumn, TableCellProperties};

// Shape elements
pub use shapes::{Transform2D, PresetGeometry, SolidFill, LineProperties, ShapeProperties, NonVisualProperties};

// DML elements
pub use dml::{Color, Fill, Outline, GradientFill, GradientStop, PatternFill, EffectExtent, Point, Size};

// Chart elements
pub use chart::{ChartKind, ChartSeries as OxmlChartSeries, ChartAxis, ChartLegend, ChartTitle, NumericData, StringData, DataPoint, CategoryPoint};
