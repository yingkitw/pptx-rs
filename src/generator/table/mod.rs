//! Table module for PPTX generation
//!
//! This module provides comprehensive table support including:
//! - Cell content and formatting
//! - Row and column management
//! - Text alignment (horizontal and vertical)
//! - Cell backgrounds and borders
//! - Text wrapping
//! - XML generation for OOXML

mod cell;
mod row;
mod builder;
mod xml;

pub use cell::{TableCell, CellAlign, CellVAlign};
pub use row::TableRow;
pub use builder::{Table, TableBuilder};
pub use xml::generate_table_xml;
