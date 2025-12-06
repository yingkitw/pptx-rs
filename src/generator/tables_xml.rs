//! Table XML generation for PPTX presentations
//!
//! Generates proper PPTX XML for tables with cells, rows, and formatting

use crate::generator::tables::{Table, TableRow, TableCell};

/// Generate table XML for a slide
pub fn generate_table_xml(table: &Table, shape_id: usize) -> String {
    let x = table.x;
    let y = table.y;
    let width = table.width();
    let height = table.height();
    let mut xml = format!(
        r#"<p:graphicFrame>
<p:nvGraphicFramePr>
<p:cNvPr id="{shape_id}" name="Table {shape_id}"/>
<p:cNvGraphicFramePr/>
<p:nvPr/>
</p:nvGraphicFramePr>
<p:xfrm>
<a:off x="{x}" y="{y}"/>
<a:ext cx="{width}" cy="{height}"/>
</p:xfrm>
<a:graphic>
<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
<a:tbl>
<a:tblPr firstRow="1" bandHVals="1"/>
<a:tblGrid>"#
    );

    // Add column widths
    for width in &table.column_widths {
        xml.push_str(&format!(r#"<a:gridCol w="{width}"/>"#));
    }

    xml.push_str("</a:tblGrid>");

    // Add rows
    for row in &table.rows {
        xml.push_str(&generate_row_xml(row));
    }

    xml.push_str(
        r#"</a:tbl>
</a:graphicData>
</a:graphic>
</p:graphicFrame>"#
    );

    xml
}

/// Generate row XML
fn generate_row_xml(row: &TableRow) -> String {
    let height = row.height.unwrap_or(400000);
    
    let mut xml = format!(r#"<a:tr h="{height}">"#);

    for cell in &row.cells {
        xml.push_str(&generate_cell_xml(cell));
    }

    xml.push_str("</a:tr>");
    xml
}

/// Generate cell XML with formatting
fn generate_cell_xml(cell: &TableCell) -> String {
    let mut xml = String::from(r#"<a:tc>"#);

    // Cell properties with margins and text anchor (vertical alignment)
    let valign = cell.valign.as_str();
    xml.push_str(&format!(
        r#"<a:tcPr marL="91440" marR="91440" marT="45720" marB="45720" anchor="{valign}" anchorCtr="0">"#
    ));

    // Background color if specified
    if let Some(color) = &cell.background_color {
        xml.push_str(&format!(
            r#"<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>"#
        ));
    } else {
        // Default light gray background for better visibility
        xml.push_str(r#"<a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill>"#);
    }

    xml.push_str("</a:tcPr>");

    // Cell text body with wrap setting and vertical alignment
    let wrap = if cell.wrap_text { "square" } else { "none" };
    let valign = cell.valign.as_str();
    let halign = cell.align.as_str();
    xml.push_str(&format!(
        r#"<a:txBody><a:bodyPr rot="0" vert="horz" anchor="{valign}" anchorCtr="0" wrap="{wrap}"/><a:lstStyle/><a:p><a:pPr algn="{halign}"/><a:r>"#
    ));

    // Build text properties with all formatting options
    let mut rpr_attrs = vec!["lang=\"en-US\"".to_string()];
    
    // Font size (default 18 points = 1800 in hundredths of a point)
    let font_size = cell.font_size.unwrap_or(18) * 100;
    rpr_attrs.push(format!("sz=\"{}\"", font_size));
    
    // Bold (always include, "0" for non-bold, "1" for bold)
    rpr_attrs.push(format!("b=\"{}\"", if cell.bold { "1" } else { "0" }));
    
    // Italic (always include, "0" for non-italic, "1" for italic)
    rpr_attrs.push(format!("i=\"{}\"", if cell.italic { "1" } else { "0" }));
    
    // dirty="0" tells PowerPoint the text doesn't need recalculation
    rpr_attrs.push("dirty=\"0\"".to_string());
    
    // Underline (only include if underlined)
    if cell.underline {
        rpr_attrs.push("u=\"sng\"".to_string());
    }
    
    // Build the rPr element with all formatting
    let mut rpr_content = String::new();
    
    // Text color if specified (default to black if not specified for visibility)
    if let Some(ref color) = cell.text_color {
        rpr_content.push_str(&format!(
            r#"<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>"#
        ));
    } else {
        // Default black text color for visibility
        rpr_content.push_str(r#"<a:solidFill><a:srgbClr val="000000"/></a:solidFill>"#);
    }
    
    // Font family - always include a default font for proper rendering
    let font_family = cell.font_family.as_deref().unwrap_or("Calibri");
    rpr_content.push_str(&format!(r#"<a:latin typeface="{font_family}"/><a:ea typeface="{font_family}"/><a:cs typeface="{font_family}"/>"#));
    
    // Build the complete rPr element (always has content now)
    xml.push_str(&format!(
        r#"<a:rPr {}>{}</a:rPr>"#,
        rpr_attrs.join(" "),
        rpr_content
    ));

    // Cell text - handle newlines by splitting into multiple paragraphs
    let text = escape_xml(&cell.text);
    if text.contains('\n') {
        // Multi-line content: split by newlines and create separate paragraphs
        let lines: Vec<&str> = text.split('\n').collect();
        for (i, line) in lines.iter().enumerate() {
            if i > 0 {
                // Close previous paragraph and start new one with same alignment
                xml.push_str(&format!(r#"</a:r></a:p><a:p><a:pPr algn="{}"/><a:r>"#, cell.align.as_str()));
                
                // Rebuild text properties for new paragraph
                let mut rpr_attrs = vec!["lang=\"en-US\"".to_string()];
                let font_size = cell.font_size.unwrap_or(20) * 100;
                rpr_attrs.push(format!("sz=\"{}\"", font_size));
                rpr_attrs.push(format!("b=\"{}\"", if cell.bold { "1" } else { "0" }));
                rpr_attrs.push(format!("i=\"{}\"", if cell.italic { "1" } else { "0" }));
                if cell.underline {
                    rpr_attrs.push("u=\"sng\"".to_string());
                }
                
                // Build the rPr element with all formatting for multi-line
                let mut rpr_content = String::new();
                
                if let Some(ref color) = cell.text_color {
                    rpr_content.push_str(&format!(
                        r#"<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>"#
                    ));
                } else {
                    rpr_content.push_str(r#"<a:solidFill><a:srgbClr val="000000"/></a:solidFill>"#);
                }
                
                let font_family = cell.font_family.as_deref().unwrap_or("Calibri");
                rpr_content.push_str(&format!(r#"<a:latin typeface="{font_family}"/><a:ea typeface="{font_family}"/><a:cs typeface="{font_family}"/>"#));
                
                xml.push_str(&format!(
                    r#"<a:rPr {}>{}</a:rPr>"#,
                    rpr_attrs.join(" "),
                    rpr_content
                ));
            }
            xml.push_str(&format!(r#"<a:t>{line}</a:t>"#));
        }
    } else {
        // Single line content
        xml.push_str(&format!(r#"<a:t>{text}</a:t>"#));
    }

    xml.push_str("</a:r></a:p></a:txBody></a:tc>");

    xml
}

/// Escape XML special characters
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

    #[test]
    fn test_generate_simple_table_xml() {
        let table = Table::from_data(
            vec![vec!["A", "B"], vec!["1", "2"]],
            vec![1000000, 1000000],
            0,
            0,
        );

        let xml = generate_table_xml(&table, 1);
        assert!(xml.contains("a:tbl"));
        assert!(xml.contains("a:tr"));
        assert!(xml.contains("a:tc"));
    }

    #[test]
    fn test_generate_cell_with_bold() {
        let cell = TableCell::new("Bold").bold();
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains(r#"b="1""#));
    }

    #[test]
    fn test_generate_cell_with_background_color() {
        let cell = TableCell::new("Colored").background_color("FF0000");
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains("FF0000"));
    }

    #[test]
    fn test_generate_cell_with_italic() {
        let cell = TableCell::new("Italic").italic();
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains(r#"i="1""#));
    }

    #[test]
    fn test_generate_cell_with_underline() {
        let cell = TableCell::new("Underline").underline();
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains(r#"u="sng""#));
    }

    #[test]
    fn test_generate_cell_with_text_color() {
        let cell = TableCell::new("Red Text").text_color("FF0000");
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains("FF0000"));
        assert!(xml.contains("srgbClr"));
    }

    #[test]
    fn test_generate_cell_with_font_size() {
        let cell = TableCell::new("Large").font_size(24);
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains("sz=\"2400\""));
    }

    #[test]
    fn test_generate_cell_with_font_family() {
        let cell = TableCell::new("Arial").font_family("Arial");
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains("typeface=\"Arial\""));
        assert!(xml.contains("latin"));
    }

    #[test]
    fn test_generate_cell_with_all_formatting() {
        let cell = TableCell::new("Styled")
            .bold()
            .italic()
            .underline()
            .text_color("0000FF")
            .background_color("FFFF00")
            .font_size(18)
            .font_family("Calibri");
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains(r#"b="1""#));
        assert!(xml.contains(r#"i="1""#));
        assert!(xml.contains(r#"u="sng""#));
        assert!(xml.contains("0000FF")); // text color
        assert!(xml.contains("FFFF00")); // background color
        assert!(xml.contains("sz=\"1800\""));
        assert!(xml.contains("typeface=\"Calibri\""));
    }

    #[test]
    fn test_escape_xml_in_cell() {
        let cell = TableCell::new("Test & <Data>");
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains("&amp;"));
        assert!(xml.contains("&lt;"));
        assert!(xml.contains("&gt;"));
    }

    #[test]
    fn test_generate_cell_with_multiline() {
        let cell = TableCell::new("Line 1\nLine 2\nLine 3");
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains("Line 1"));
        assert!(xml.contains("Line 2"));
        assert!(xml.contains("Line 3"));
        // Should have multiple paragraphs
        assert!(xml.matches("</a:p>").count() >= 3);
    }
}
