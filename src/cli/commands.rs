//! CLI commands implementation

use std::fs;
use std::path::PathBuf;
use crate::generator::{self, SlideContent};

pub struct CreateCommand;
pub struct FromMarkdownCommand;
pub struct InfoCommand;

impl CreateCommand {
    pub fn execute(
        output: &str,
        title: Option<&str>,
        slides: usize,
        _template: Option<&str>,
    ) -> Result<(), String> {
        // Create output directory if needed
        if let Some(parent) = PathBuf::from(output).parent() {
            if !parent.as_os_str().is_empty() {
                fs::create_dir_all(parent)
                    .map_err(|e| format!("Failed to create directory: {e}"))?;
            }
        }

        let title = title.unwrap_or("Presentation");

        // Generate proper PPTX file
        let pptx_data = generator::create_pptx(title, slides)
            .map_err(|e| format!("Failed to generate PPTX: {e}"))?;

        // Write to file
        fs::write(output, pptx_data)
            .map_err(|e| format!("Failed to write file: {e}"))?;

        Ok(())
    }
}

impl FromMarkdownCommand {
    pub fn execute(
        input: &str,
        output: &str,
        title: Option<&str>,
    ) -> Result<(), String> {
        // Read markdown file
        let md_content = fs::read_to_string(input)
            .map_err(|e| format!("Failed to read markdown file: {e}"))?;

        // Parse markdown into slides
        let slides = Self::parse_markdown(&md_content)?;

        if slides.is_empty() {
            return Err("No slides found in markdown file".to_string());
        }

        // Create output directory if needed
        if let Some(parent) = PathBuf::from(output).parent() {
            if !parent.as_os_str().is_empty() {
                fs::create_dir_all(parent)
                    .map_err(|e| format!("Failed to create directory: {e}"))?;
            }
        }

        let title = title.unwrap_or("Presentation from Markdown");

        // Generate PPTX with content
        let pptx_data = generator::create_pptx_with_content(title, slides)
            .map_err(|e| format!("Failed to generate PPTX: {e}"))?;

        // Write to file
        fs::write(output, pptx_data)
            .map_err(|e| format!("Failed to write file: {e}"))?;

        Ok(())
    }

    fn parse_markdown(content: &str) -> Result<Vec<SlideContent>, String> {
        let mut slides = Vec::new();
        let mut current_slide: Option<SlideContent> = None;

        for line in content.lines() {
            let trimmed = line.trim();

            // Check for slide title (# heading)
            if trimmed.starts_with("# ") {
                // Save previous slide if exists
                if let Some(slide) = current_slide.take() {
                    slides.push(slide);
                }

                // Create new slide
                let title = trimmed[2..].trim().to_string();
                current_slide = Some(SlideContent::new(&title));
            }
            // Check for bullet points (- or * or +)
            else if trimmed.starts_with("- ") || trimmed.starts_with("* ") || trimmed.starts_with("+ ") {
                if let Some(ref mut slide) = current_slide {
                    let bullet = trimmed[2..].trim().to_string();
                    if !bullet.is_empty() {
                        *slide = slide.clone().add_bullet(&bullet);
                    }
                } else {
                    // Create a default slide if no title exists
                    let bullet = trimmed[2..].trim().to_string();
                    if !bullet.is_empty() {
                        let mut slide = SlideContent::new("Slide");
                        slide = slide.add_bullet(&bullet);
                        current_slide = Some(slide);
                    }
                }
            }
            // Skip empty lines and other content
        }

        // Add last slide if exists
        if let Some(slide) = current_slide {
            slides.push(slide);
        }

        Ok(slides)
    }
}

impl InfoCommand {
    pub fn execute(file: &str) -> Result<(), String> {
        let metadata = fs::metadata(file)
            .map_err(|e| format!("File not found: {e}"))?;

        let size = metadata.len();
        let modified = metadata
            .modified()
            .ok()
            .and_then(|t| t.elapsed().ok())
            .map(|d| format!("{d:?} ago"))
            .unwrap_or_else(|| "unknown".to_string());

        println!("File Information");
        println!("================");
        println!("Path:     {file}");
        println!("Size:     {size} bytes");
        println!("Modified: {modified}");
        let is_file = metadata.is_file();
        println!("Is file:  {is_file}");

        // Try to read and parse as XML
        if let Ok(content) = fs::read_to_string(file) {
            if content.starts_with("<?xml") {
                println!("\nPresentation Information");
                println!("========================");
                if let Some(title_start) = content.find("<title>") {
                    if let Some(title_end) = content[title_start + 7..].find("</title>") {
                        let title = &content[title_start + 7..title_start + 7 + title_end];
                        println!("Title: {title}");
                    }
                }
                if let Some(slides_start) = content.find("count=\"") {
                    let search_from = slides_start + 7;
                    if let Some(slides_end) = content[search_from..].find("\"") {
                        let count_str = &content[search_from..search_from + slides_end];
                        println!("Slides: {count_str}");
                    }
                }
            }
        }

        Ok(())
    }
}

#[allow(dead_code)]
fn escape_xml(s: &str) -> String {
    s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\"", "&quot;")
        .replace("'", "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use std::path::Path;

    #[test]
    fn test_create_command() {
        let output = "/tmp/test_presentation.pptx";
        let result = CreateCommand::execute(output, Some("Test"), 3, None);
        assert!(result.is_ok());
        assert!(Path::new(output).exists());
        
        // Cleanup
        let _ = fs::remove_file(output);
    }

    #[test]
    fn test_escape_xml() {
        assert_eq!(escape_xml("a & b"), "a &amp; b");
        assert_eq!(escape_xml("<tag>"), "&lt;tag&gt;");
        assert_eq!(escape_xml("\"quoted\""), "&quot;quoted&quot;");
    }
}
