//! Package parts module
//!
//! Provides abstraction for PPTX package parts (files within the ZIP).
//! Each part type handles its own XML generation and parsing.

pub mod base;
pub mod presentation;
pub mod slide;
pub mod image;
pub mod chart;
pub mod coreprops;
pub mod relationships;

// Re-export main types
pub use base::{Part, PartType, ContentType};
pub use presentation::PresentationPart;
pub use slide::SlidePart;
pub use image::ImagePart;
pub use chart::ChartPart;
pub use coreprops::CorePropertiesPart;
pub use relationships::{Relationship, RelationshipType, Relationships};
