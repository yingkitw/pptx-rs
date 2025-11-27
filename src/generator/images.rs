//! Image handling for PPTX presentations
//!
//! Handles image metadata, embedding, and XML generation

use std::path::Path;

/// Image metadata and properties
#[derive(Clone, Debug)]
pub struct Image {
    pub filename: String,
    pub width: u32,      // in EMU
    pub height: u32,     // in EMU
    pub x: u32,          // Position X in EMU
    pub y: u32,          // Position Y in EMU
    pub format: String,  // PNG, JPG, GIF, etc.
}

impl Image {
    /// Create a new image
    pub fn new(filename: &str, width: u32, height: u32, format: &str) -> Self {
        Image {
            filename: filename.to_string(),
            width,
            height,
            x: 0,
            y: 0,
            format: format.to_uppercase(),
        }
    }

    /// Set image position
    pub fn position(mut self, x: u32, y: u32) -> Self {
        self.x = x;
        self.y = y;
        self
    }

    /// Get aspect ratio
    pub fn aspect_ratio(&self) -> f64 {
        self.width as f64 / self.height as f64
    }

    /// Scale image to width while maintaining aspect ratio
    pub fn scale_to_width(mut self, width: u32) -> Self {
        let ratio = self.aspect_ratio();
        self.width = width;
        self.height = (width as f64 / ratio) as u32;
        self
    }

    /// Scale image to height while maintaining aspect ratio
    pub fn scale_to_height(mut self, height: u32) -> Self {
        let ratio = self.aspect_ratio();
        self.height = height;
        self.width = (height as f64 * ratio) as u32;
        self
    }

    /// Get file extension from filename
    pub fn extension(&self) -> String {
        Path::new(&self.filename)
            .extension()
            .and_then(|ext| ext.to_str())
            .map(|s| s.to_lowercase())
            .unwrap_or_else(|| self.format.to_lowercase())
    }

    /// Get MIME type for the image format
    pub fn mime_type(&self) -> String {
        match self.format.as_str() {
            "PNG" => "image/png".to_string(),
            "JPG" | "JPEG" => "image/jpeg".to_string(),
            "GIF" => "image/gif".to_string(),
            "BMP" => "image/bmp".to_string(),
            "TIFF" => "image/tiff".to_string(),
            "SVG" => "image/svg+xml".to_string(),
            _ => "application/octet-stream".to_string(),
        }
    }
}

/// Image builder for fluent API
pub struct ImageBuilder {
    filename: String,
    width: u32,
    height: u32,
    x: u32,
    y: u32,
    format: String,
}

impl ImageBuilder {
    /// Create a new image builder
    pub fn new(filename: &str, width: u32, height: u32) -> Self {
        let format = Path::new(filename)
            .extension()
            .and_then(|ext| ext.to_str())
            .map(|s| s.to_uppercase())
            .unwrap_or_else(|| "PNG".to_string());

        ImageBuilder {
            filename: filename.to_string(),
            width,
            height,
            x: 0,
            y: 0,
            format,
        }
    }

    /// Set image position
    pub fn position(mut self, x: u32, y: u32) -> Self {
        self.x = x;
        self.y = y;
        self
    }

    /// Set image format
    pub fn format(mut self, format: &str) -> Self {
        self.format = format.to_uppercase();
        self
    }

    /// Scale to width
    pub fn scale_to_width(mut self, width: u32) -> Self {
        let ratio = self.width as f64 / self.height as f64;
        self.width = width;
        self.height = (width as f64 / ratio) as u32;
        self
    }

    /// Scale to height
    pub fn scale_to_height(mut self, height: u32) -> Self {
        let ratio = self.width as f64 / self.height as f64;
        self.height = height;
        self.width = (height as f64 * ratio) as u32;
        self
    }

    /// Build the image
    pub fn build(self) -> Image {
        Image {
            filename: self.filename,
            width: self.width,
            height: self.height,
            x: self.x,
            y: self.y,
            format: self.format,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_image_creation() {
        let img = Image::new("test.png", 1920, 1080, "PNG");
        assert_eq!(img.filename, "test.png");
        assert_eq!(img.width, 1920);
        assert_eq!(img.height, 1080);
    }

    #[test]
    fn test_image_position() {
        let img = Image::new("test.png", 1920, 1080, "PNG")
            .position(500000, 1000000);
        assert_eq!(img.x, 500000);
        assert_eq!(img.y, 1000000);
    }

    #[test]
    fn test_image_aspect_ratio() {
        let img = Image::new("test.png", 1920, 1080, "PNG");
        let ratio = img.aspect_ratio();
        assert!((ratio - 1.777).abs() < 0.01);
    }

    #[test]
    fn test_image_scale_to_width() {
        let img = Image::new("test.png", 1920, 1080, "PNG")
            .scale_to_width(960);
        assert_eq!(img.width, 960);
        assert_eq!(img.height, 540);
    }

    #[test]
    fn test_image_scale_to_height() {
        let img = Image::new("test.png", 1920, 1080, "PNG")
            .scale_to_height(540);
        assert_eq!(img.width, 960);
        assert_eq!(img.height, 540);
    }

    #[test]
    fn test_image_extension() {
        let img = Image::new("photo.jpg", 1920, 1080, "JPEG");
        assert_eq!(img.extension(), "jpg");
    }

    #[test]
    fn test_image_mime_types() {
        assert_eq!(
            Image::new("test.png", 100, 100, "PNG").mime_type(),
            "image/png"
        );
        assert_eq!(
            Image::new("test.jpg", 100, 100, "JPG").mime_type(),
            "image/jpeg"
        );
        assert_eq!(
            Image::new("test.gif", 100, 100, "GIF").mime_type(),
            "image/gif"
        );
    }

    #[test]
    fn test_image_builder() {
        let img = ImageBuilder::new("photo.png", 1920, 1080)
            .position(500000, 1000000)
            .scale_to_width(960)
            .build();

        assert_eq!(img.filename, "photo.png");
        assert_eq!(img.width, 960);
        assert_eq!(img.height, 540);
        assert_eq!(img.x, 500000);
        assert_eq!(img.y, 1000000);
    }

    #[test]
    fn test_image_builder_auto_format() {
        let img = ImageBuilder::new("photo.jpg", 1920, 1080).build();
        assert_eq!(img.format, "JPG");
    }
}
