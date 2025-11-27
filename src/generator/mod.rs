//! PPTX file generator - creates proper ZIP-based PPTX files

pub mod builder;
pub mod xml;
pub mod text;
pub mod shapes;
pub mod tables;
pub mod tables_xml;
pub mod images;
pub mod images_xml;

pub use builder::{create_pptx, create_pptx_with_content};
pub use xml::SlideContent;
pub use text::{TextFormat, FormattedText};
pub use shapes::{Shape, ShapeType, ShapeFill, ShapeLine};
pub use tables::{Table, TableRow, TableCell, TableBuilder};
pub use images::{Image, ImageBuilder};
pub use images_xml::{generate_image_xml, generate_image_relationship, generate_image_content_type};

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_slide_content_builder() {
        let slide = SlideContent::new("Title")
            .add_bullet("Point 1")
            .add_bullet("Point 2");
        
        assert_eq!(slide.title, "Title");
        assert_eq!(slide.content.len(), 2);
        assert_eq!(slide.content[0], "Point 1");
    }
}
