//! ER diagram parsing and rendering

use std::collections::HashMap;
use crate::generator::{Shape, ShapeType, ShapeFill, ShapeLine};
use crate::generator::connectors::{Connector, ConnectorType, ConnectorLine, ArrowType};
use super::types::DiagramElements;

/// Generate shapes and connectors for an ER diagram
pub fn generate_elements(code: &str) -> DiagramElements {
    let mut shapes = Vec::new();
    let mut connectors = Vec::new();
    
    let mut entities: HashMap<String, Vec<String>> = HashMap::new();
    let mut relationships: Vec<(String, String, String)> = Vec::new();
    let mut current_entity = String::new();
    
    for line in code.lines().skip(1) {
        let line = line.trim();
        if line.is_empty() || line.starts_with("%%") {
            continue;
        }
        
        if line.contains("||") || line.contains("}|") || line.contains("|{") || line.contains("o{") {
            let parts: Vec<&str> = line.split(|c: char| c == '|' || c == '{' || c == '}' || c == 'o' || c == '-').collect();
            let parts: Vec<&str> = parts.into_iter().filter(|s| !s.is_empty() && !s.contains(':') && s.chars().any(|c| c.is_alphabetic())).collect();
            if parts.len() >= 2 {
                let e1 = parts[0].trim().to_string();
                let e2 = parts[parts.len()-1].trim().to_string();
                if !entities.contains_key(&e1) { entities.insert(e1.clone(), Vec::new()); }
                if !entities.contains_key(&e2) { entities.insert(e2.clone(), Vec::new()); }
                relationships.push((e1, e2, "relates".to_string()));
            }
        }
        else if line.contains('{') {
            current_entity = line.split('{').next().unwrap_or("").trim().to_string();
            if !entities.contains_key(&current_entity) {
                entities.insert(current_entity.clone(), Vec::new());
            }
        } else if line == "}" {
            current_entity.clear();
        } else if !current_entity.is_empty() && !line.is_empty() {
            if let Some(attrs) = entities.get_mut(&current_entity) {
                attrs.push(line.to_string());
            }
        }
    }
    
    // Layout parameters
    let start_x = 500_000u32;
    let start_y = 1_600_000u32;
    let entity_width = 2_200_000u32;
    let header_height = 400_000u32;
    let attr_height = 280_000u32;
    let h_spacing = 2_800_000u32;
    let v_spacing = 2_500_000u32;
    
    let mut entity_positions: HashMap<String, (u32, u32)> = HashMap::new();
    
    for (i, (entity_name, attrs)) in entities.iter().enumerate() {
        let x = start_x + (i as u32 % 3) * h_spacing;
        let y = start_y + (i as u32 / 3) * v_spacing;
        entity_positions.insert(entity_name.clone(), (x, y));
        
        let header = Shape::new(ShapeType::Rectangle, x, y, entity_width, header_height)
            .with_fill(ShapeFill::new("C2185B"))
            .with_line(ShapeLine::new("880E4F", 2))
            .with_text(entity_name);
        shapes.push(header);
        
        let attrs_text = attrs.join("\n");
        let attrs_box_height = (attrs.len().max(1) as u32) * attr_height;
        let attrs_shape = Shape::new(ShapeType::Rectangle, x, y + header_height, entity_width, attrs_box_height)
            .with_fill(ShapeFill::new("FCE4EC"))
            .with_line(ShapeLine::new("880E4F", 1))
            .with_text(&attrs_text);
        shapes.push(attrs_shape);
    }
    
    for (e1, e2, _) in &relationships {
        if let (Some(&(x1, y1)), Some(&(x2, y2))) = 
            (entity_positions.get(e1), entity_positions.get(e2)) 
        {
            let connector = Connector::new(
                ConnectorType::Elbow,
                x1 + entity_width, y1 + header_height / 2,
                x2, y2 + header_height / 2
            )
            .with_line(ConnectorLine::new("880E4F", 19050))
            .with_end_arrow(ArrowType::Diamond);
            connectors.push(connector);
        }
    }
    
    DiagramElements { shapes, connectors }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_generate_er_diagram() {
        // Test with entity definitions
        let code = "erDiagram\n    CUSTOMER {\n        string name\n    }\n    ORDER {\n        int id\n    }\n    CUSTOMER ||--o{ ORDER : places";
        let elements = generate_elements(code);
        assert!(!elements.shapes.is_empty());
    }

    #[test]
    fn test_er_diagram_empty() {
        // Empty ER diagram should not panic
        let code = "erDiagram";
        let elements = generate_elements(code);
        assert!(elements.shapes.is_empty());
    }
}
