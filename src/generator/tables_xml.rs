//! Table XML generation for PPTX presentations
//!
//! Generates proper PPTX XML for tables with cells, rows, and formatting

use crate::generator::tables::{Table, TableRow, TableCell};

/// Generate table XML for a slide
pub fn generate_table_xml(table: &Table, shape_id: usize) -> String {
    let mut xml = format!(
        r#"<p:graphicFrame>
<p:nvGraphicFramePr>
<p:cNvPr id="{}" name="Table {}"/>
<p:cNvGraphicFramePr/>
<p:nvPr/>
</p:nvGraphicFramePr>
<p:xfrm>
<a:off x="{}" y="{}"/>
<a:ext cx="{}" cy="{}"/>
</p:xfrm>
<a:graphic>
<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
<a:tbl>
<a:tblPr firstRow="1" bandHVals="1"/>
<a:tblGrid>"#,
        shape_id,
        shape_id,
        table.x,
        table.y,
        table.width(),
        table.height()
    );

    // Add column widths
    for width in &table.column_widths {
        xml.push_str(&format!(r#"<a:gridCol w="{}"/>"#, width));
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
    
    let mut xml = format!(r#"<a:tr h="{}">"#, height);

    for cell in &row.cells {
        xml.push_str(&generate_cell_xml(cell));
    }

    xml.push_str("</a:tr>");
    xml
}

/// Generate cell XML with formatting
fn generate_cell_xml(cell: &TableCell) -> String {
    let mut xml = String::from(r#"<a:tc>"#);

    // Cell properties
    xml.push_str("<a:tcPr>");

    // Background color if specified
    if let Some(color) = &cell.background_color {
        xml.push_str(&format!(
            r#"<a:solidFill><a:srgbClr val="{}"/></a:solidFill>"#,
            color
        ));
    }

    xml.push_str("</a:tcPr>");

    // Cell text body
    xml.push_str(r#"<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>"#);

    // Text properties
    let bold = if cell.bold { "1" } else { "0" };
    xml.push_str(&format!(
        r#"<a:rPr lang="en-US" sz="2400" b="{}"/>"#,
        bold
    ));

    // Cell text
    xml.push_str(&format!(r#"<a:t>{}</a:t>"#, escape_xml(&cell.text)));

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
    fn test_generate_cell_with_color() {
        let cell = TableCell::new("Colored").background_color("FF0000");
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains("FF0000"));
    }

    #[test]
    fn test_escape_xml_in_cell() {
        let cell = TableCell::new("Test & <Data>");
        let xml = generate_cell_xml(&cell);
        assert!(xml.contains("&amp;"));
        assert!(xml.contains("&lt;"));
        assert!(xml.contains("&gt;"));
    }
}
