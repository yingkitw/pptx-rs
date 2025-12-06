//! XML generation for tables in PPTX format

use super::builder::Table;
use super::row::TableRow;
use super::cell::TableCell;

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
<a:tblPr firstRow="1" bandRow="1"/>
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

    // Cell properties with margins and vertical alignment
    let valign = cell.valign.as_str();
    xml.push_str(&format!(
        r#"<a:tcPr marL="91440" marR="91440" marT="45720" marB="45720" anchor="{valign}">"#
    ));

    // Background color
    if let Some(color) = &cell.background_color {
        xml.push_str(&format!(
            r#"<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>"#
        ));
    } else {
        // Default light gray background
        xml.push_str(r#"<a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill>"#);
    }

    xml.push_str("</a:tcPr>");

    // Text body with alignment and wrapping
    let wrap = if cell.wrap_text { "square" } else { "none" };
    let halign = cell.align.as_str();
    
    xml.push_str(&format!(
        r#"<a:txBody><a:bodyPr wrap="{wrap}" anchor="{valign}"/><a:lstStyle/><a:p><a:pPr algn="{halign}"/>"#
    ));

    // Text run with formatting
    xml.push_str(&generate_text_run(cell));

    xml.push_str("</a:p></a:txBody></a:tc>");

    xml
}

/// Generate text run with formatting
fn generate_text_run(cell: &TableCell) -> String {
    let mut xml = String::from("<a:r>");

    // Run properties
    let font_size = cell.font_size.unwrap_or(18) * 100;
    let bold = if cell.bold { "1" } else { "0" };
    let italic = if cell.italic { "1" } else { "0" };
    
    xml.push_str(&format!(
        r#"<a:rPr lang="en-US" sz="{font_size}" b="{bold}" i="{italic}" dirty="0">"#
    ));

    // Text color
    if let Some(ref color) = cell.text_color {
        xml.push_str(&format!(
            r#"<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>"#
        ));
    } else {
        // Default black text
        xml.push_str(r#"<a:solidFill><a:srgbClr val="000000"/></a:solidFill>"#);
    }

    // Font family (always include for proper rendering)
    let font_family = cell.font_family.as_deref().unwrap_or("Calibri");
    xml.push_str(&format!(
        r#"<a:latin typeface="{font_family}"/><a:ea typeface="{font_family}"/><a:cs typeface="{font_family}"/>"#
    ));

    // Underline
    if cell.underline {
        xml.push_str(r#"<a:uFill><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:uFill>"#);
    }

    xml.push_str("</a:rPr>");

    // Text content
    let text = escape_xml(&cell.text);
    xml.push_str(&format!(r#"<a:t>{text}</a:t>"#));

    xml.push_str("</a:r>");
    xml
}

/// Escape XML special characters
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
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
        assert!(xml.contains("<a:t>A</a:t>"));
        assert!(xml.contains("<a:t>B</a:t>"));
    }

    #[test]
    fn test_generate_cell_with_bold() {
        let cell = TableCell::new("Bold").bold();
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains(r#"b="1""#));
    }

    #[test]
    fn test_generate_cell_with_text_color() {
        let cell = TableCell::new("Red").text_color("FF0000");
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains("FF0000"));
    }

    #[test]
    fn test_generate_cell_with_background_color() {
        let cell = TableCell::new("Blue BG").background_color("0000FF");
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains("0000FF"));
    }

    #[test]
    fn test_generate_cell_with_alignment() {
        let cell = TableCell::new("Left").align_left();
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains(r#"algn="l""#));
    }

    #[test]
    fn test_escape_xml() {
        let cell = TableCell::new("Test & <Data>");
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains("&amp;"));
        assert!(xml.contains("&lt;"));
        assert!(xml.contains("&gt;"));
    }

    #[test]
    fn test_font_always_included() {
        let cell = TableCell::new("Test");
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains(r#"<a:latin typeface="Calibri"/>"#));
        assert!(xml.contains(r#"<a:ea typeface="Calibri"/>"#));
        assert!(xml.contains(r#"<a:cs typeface="Calibri"/>"#));
    }
}
