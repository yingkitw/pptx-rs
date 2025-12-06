//! Table row definition

use super::cell::TableCell;

/// Table row containing cells
#[derive(Clone, Debug)]
pub struct TableRow {
    /// Cells in this row
    pub cells: Vec<TableCell>,
    /// Row height in EMU (optional)
    pub height: Option<u32>,
}

impl TableRow {
    /// Create a new table row with cells
    pub fn new(cells: Vec<TableCell>) -> Self {
        TableRow {
            cells,
            height: None,
        }
    }

    /// Set row height in EMU
    pub fn with_height(mut self, height: u32) -> Self {
        self.height = Some(height);
        self
    }

    /// Get the number of cells in this row
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_row_new() {
        let row = TableRow::new(vec![
            TableCell::new("A"),
            TableCell::new("B"),
        ]);
        assert_eq!(row.cell_count(), 2);
        assert!(row.height.is_none());
    }

    #[test]
    fn test_row_with_height() {
        let row = TableRow::new(vec![TableCell::new("A")])
            .with_height(500000);
        assert_eq!(row.height, Some(500000));
    }
}
