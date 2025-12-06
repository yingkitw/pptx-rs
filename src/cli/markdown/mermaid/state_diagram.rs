//! State diagram parsing and rendering

use std::collections::HashMap;
use crate::generator::{Shape, ShapeType, ShapeFill, ShapeLine};
use crate::generator::connectors::{Connector, ConnectorType, ConnectorLine, ArrowType};
use super::types::DiagramElements;

/// Generate shapes and connectors for a state diagram
pub fn generate_elements(code: &str) -> DiagramElements {
    let mut shapes = Vec::new();
    let mut connectors = Vec::new();
    
    let mut states: Vec<String> = Vec::new();
    let mut transitions: Vec<(String, String, String)> = Vec::new();
    
    for line in code.lines().skip(1) {
        let line = line.trim();
        if line.is_empty() || line.starts_with("%%") || line.starts_with("direction") {
            continue;
        }
        
        if line.contains("-->") {
            let parts: Vec<&str> = line.split("-->").collect();
            if parts.len() >= 2 {
                let from = parts[0].trim().to_string();
                let (to, label) = if let Some((t, l)) = parts[1].split_once(':') {
                    (t.trim().to_string(), l.trim().to_string())
                } else {
                    (parts[1].trim().to_string(), String::new())
                };
                
                let from_state = if from == "[*]" { "Start".to_string() } else { from };
                let to_state = if to == "[*]" { "End".to_string() } else { to };
                
                if !states.contains(&from_state) { states.push(from_state.clone()); }
                if !states.contains(&to_state) { states.push(to_state.clone()); }
                
                transitions.push((from_state, to_state, label));
            }
        }
    }
    
    // Layout parameters
    let start_x = 1_000_000u32;
    let start_y = 1_800_000u32;
    let state_width = 1_500_000u32;
    let state_height = 500_000u32;
    let h_spacing = 2_200_000u32;
    let v_spacing = 1_200_000u32;
    
    let mut state_positions: HashMap<String, (u32, u32)> = HashMap::new();
    
    for (i, state) in states.iter().enumerate() {
        let x = start_x + (i as u32 % 3) * h_spacing;
        let y = start_y + (i as u32 / 3) * v_spacing;
        state_positions.insert(state.clone(), (x, y));
        
        let shape_type = if state == "Start" || state == "End" {
            ShapeType::Ellipse
        } else {
            ShapeType::RoundedRectangle
        };
        
        let fill_color = if state == "Start" { "000000" }
                        else if state == "End" { "000000" }
                        else { "E0F7FA" };
        
        let shape = Shape::new(shape_type, x, y, state_width, state_height)
            .with_fill(ShapeFill::new(fill_color))
            .with_line(ShapeLine::new("00838F", 2))
            .with_text(state);
        shapes.push(shape);
    }
    
    for (from, to, label) in &transitions {
        if let (Some(&(from_x, from_y)), Some(&(to_x, to_y))) = 
            (state_positions.get(from), state_positions.get(to)) 
        {
            let mut connector = Connector::new(
                ConnectorType::Elbow,
                from_x + state_width, from_y + state_height / 2,
                to_x, to_y + state_height / 2
            )
            .with_line(ConnectorLine::new("00838F", 19050))
            .with_end_arrow(ArrowType::Triangle);
            
            if !label.is_empty() {
                connector = connector.with_label(label);
            }
            connectors.push(connector);
        }
    }
    
    DiagramElements { shapes, connectors }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_generate_state_diagram() {
        let code = "stateDiagram\n    [*] --> Active\n    Active --> [*]";
        let elements = generate_elements(code);
        assert!(!elements.shapes.is_empty());
    }
}
