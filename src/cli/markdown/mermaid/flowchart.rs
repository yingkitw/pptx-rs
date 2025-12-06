//! Flowchart diagram parsing and rendering

use std::collections::HashMap;
use crate::generator::{Shape, ShapeType, ShapeFill, ShapeLine};
use crate::generator::connectors::{Connector, ConnectorType, ConnectorLine, ArrowType, LineDash};
use super::types::*;

/// Parse flowchart direction from first line
fn parse_direction(first_line: &str) -> FlowDirection {
    let line = first_line.to_uppercase();
    if line.contains("LR") {
        FlowDirection::LeftToRight
    } else if line.contains("RL") {
        FlowDirection::RightToLeft
    } else if line.contains("BT") {
        FlowDirection::BottomToTop
    } else {
        FlowDirection::TopToBottom
    }
}

/// Parse a flowchart from Mermaid code
pub fn parse(code: &str) -> Flowchart {
    let mut lines = code.lines();
    let first_line = lines.next().unwrap_or("");
    let direction = parse_direction(first_line);
    
    let mut nodes: HashMap<String, FlowNode> = HashMap::new();
    let mut connections: Vec<FlowConnection> = Vec::new();
    let mut subgraphs: Vec<Subgraph> = Vec::new();
    let mut current_subgraph: Option<Subgraph> = None;
    
    for line in lines {
        let line = line.trim();
        if line.is_empty() || line.starts_with("%%") {
            continue;
        }
        
        // Handle subgraph start
        if line.starts_with("subgraph") {
            let name = line.strip_prefix("subgraph").unwrap_or("").trim().to_string();
            current_subgraph = Some(Subgraph { name, nodes: Vec::new() });
            continue;
        }
        
        // Handle subgraph end
        if line == "end" {
            if let Some(sg) = current_subgraph.take() {
                subgraphs.push(sg);
            }
            continue;
        }
        
        // Parse connections: A --> B, A --> B[Label], A[Text] --> B[Text]
        if let Some((from_part, rest)) = split_connection(line) {
            let (arrow_type, to_part) = parse_arrow_and_rest(&rest);
            
            // Parse from node
            let (from_id, from_node) = parse_node_def(&from_part);
            if let Some(node) = from_node {
                nodes.entry(from_id.clone()).or_insert(node);
                if let Some(ref mut sg) = current_subgraph {
                    if !sg.nodes.contains(&from_id) {
                        sg.nodes.push(from_id.clone());
                    }
                }
            }
            
            // Parse to node (may have label on arrow)
            let (to_part_clean, arrow_label) = extract_arrow_label(&to_part);
            let (to_id, to_node) = parse_node_def(&to_part_clean);
            if let Some(node) = to_node {
                nodes.entry(to_id.clone()).or_insert(node);
                if let Some(ref mut sg) = current_subgraph {
                    if !sg.nodes.contains(&to_id) {
                        sg.nodes.push(to_id.clone());
                    }
                }
            }
            
            connections.push(FlowConnection {
                from: from_id,
                to: to_id,
                label: arrow_label,
                arrow_type,
            });
        } else {
            // Standalone node definition
            let (id, node) = parse_node_def(line);
            if let Some(n) = node {
                nodes.entry(id.clone()).or_insert(n);
                if let Some(ref mut sg) = current_subgraph {
                    if !sg.nodes.contains(&id) {
                        sg.nodes.push(id);
                    }
                }
            }
        }
    }
    
    Flowchart {
        direction,
        nodes: nodes.into_values().collect(),
        connections,
        subgraphs,
    }
}

/// Split line at connection arrow
fn split_connection(line: &str) -> Option<(String, String)> {
    for arrow in ["==>", "-.->", "-->", "---", "->"] {
        if let Some(pos) = line.find(arrow) {
            let from = line[..pos].trim().to_string();
            let rest = line[pos..].to_string();
            return Some((from, rest));
        }
    }
    None
}

/// Parse arrow type and get the rest of the string
fn parse_arrow_and_rest(s: &str) -> (ArrowStyle, String) {
    if s.starts_with("==>") {
        (ArrowStyle::Thick, s[3..].trim().to_string())
    } else if s.starts_with("-.->") {
        (ArrowStyle::Dotted, s[4..].trim().to_string())
    } else if s.starts_with("-->") {
        (ArrowStyle::Arrow, s[3..].trim().to_string())
    } else if s.starts_with("---") {
        (ArrowStyle::Open, s[3..].trim().to_string())
    } else if s.starts_with("->") {
        (ArrowStyle::Arrow, s[2..].trim().to_string())
    } else {
        (ArrowStyle::Arrow, s.to_string())
    }
}

/// Extract arrow label like |text|
fn extract_arrow_label(s: &str) -> (String, Option<String>) {
    if let Some(start) = s.find('|') {
        if let Some(end) = s[start+1..].find('|') {
            let label = s[start+1..start+1+end].to_string();
            let rest = s[start+2+end..].trim().to_string();
            return (rest, Some(label));
        }
    }
    (s.to_string(), None)
}

/// Parse a node definition like A[Text] or B(Text) or C{Text}
fn parse_node_def(s: &str) -> (String, Option<FlowNode>) {
    let s = s.trim();
    
    // Try different bracket types
    for (open, close, shape) in [
        ("((", "))", NodeShape::Circle),
        ("([", "])", NodeShape::Stadium),
        ("{{", "}}", NodeShape::Hexagon),
        ("[", "]", NodeShape::Rectangle),
        ("(", ")", NodeShape::RoundedRect),
        ("{", "}", NodeShape::Diamond),
    ] {
        if let Some(start) = s.find(open) {
            let id = s[..start].trim().to_string();
            if let Some(end) = s[start+open.len()..].find(close) {
                let label = s[start+open.len()..start+open.len()+end].to_string();
                return (id.clone(), Some(FlowNode { id, label, shape }));
            }
        }
    }
    
    // Plain node ID without brackets
    let id = s.to_string();
    if !id.is_empty() && id.chars().all(|c| c.is_alphanumeric() || c == '_') {
        return (id.clone(), Some(FlowNode { 
            id: id.clone(), 
            label: id, 
            shape: NodeShape::Rectangle 
        }));
    }
    
    (s.to_string(), None)
}

/// Generate shapes and connectors for a flowchart
pub fn generate_elements(flowchart: &Flowchart) -> DiagramElements {
    let mut shapes = Vec::new();
    let mut connectors = Vec::new();
    let node_count = flowchart.nodes.len();
    
    if node_count == 0 {
        return DiagramElements { shapes, connectors };
    }
    
    // Layout parameters (in EMUs)
    let node_width = 1_400_000u32;
    let node_height = 500_000u32;
    let h_spacing = 1_800_000u32;
    let v_spacing = 900_000u32;
    
    let is_horizontal = matches!(flowchart.direction, FlowDirection::LeftToRight | FlowDirection::RightToLeft);
    
    let mut node_positions: HashMap<String, (u32, u32)> = HashMap::new();
    let mut _shape_id = 10u32;
    
    // If we have subgraphs, layout by subgraph
    if !flowchart.subgraphs.is_empty() {
        let mut subgraph_x = 500_000u32;
        let subgraph_start_y = 1_600_000u32;
        
        for (sg_idx, subgraph) in flowchart.subgraphs.iter().enumerate() {
            let sg_width = node_width + 400_000;
            let sg_height = (subgraph.nodes.len() as u32) * v_spacing + 400_000;
            let sg_x = subgraph_x;
            let sg_y = subgraph_start_y;
            
            // Subgraph background shape
            let sg_shape = Shape::new(ShapeType::RoundedRectangle, sg_x, sg_y, sg_width, sg_height)
                .with_fill(ShapeFill::new(get_subgraph_color(sg_idx)))
                .with_line(ShapeLine::new("757575", 1))
                .with_text(&subgraph.name);
            shapes.push(sg_shape);
            
            // Layout nodes within subgraph
            for (node_idx, node_id) in subgraph.nodes.iter().enumerate() {
                if let Some(node) = flowchart.nodes.iter().find(|n| &n.id == node_id) {
                    let x = sg_x + 200_000;
                    let y = sg_y + 300_000 + (node_idx as u32) * v_spacing;
                    
                    node_positions.insert(node.id.clone(), (x, y));
                    
                    let shape = create_node_shape(node, x, y, node_width, node_height);
                    shapes.push(shape);
                    _shape_id += 1;
                }
            }
            
            subgraph_x += sg_width + 600_000;
        }
        
        // Layout any nodes not in subgraphs
        let mut orphan_y = subgraph_start_y;
        for node in &flowchart.nodes {
            if !node_positions.contains_key(&node.id) {
                let x = subgraph_x;
                let y = orphan_y;
                
                node_positions.insert(node.id.clone(), (x, y));
                
                let shape = create_node_shape(node, x, y, node_width, node_height);
                shapes.push(shape);
                _shape_id += 1;
                
                orphan_y += v_spacing;
            }
        }
    } else {
        // Simple grid layout without subgraphs
        let start_x = 1_000_000u32;
        let start_y = 1_800_000u32;
        let cols = if is_horizontal { node_count.min(5) } else { 1 };
        
        for (i, node) in flowchart.nodes.iter().enumerate() {
            let col = i % cols;
            let row = i / cols;
            
            let (x, y) = if is_horizontal {
                (start_x + (col as u32) * h_spacing, start_y + (row as u32) * v_spacing)
            } else {
                (start_x + (col as u32) * h_spacing, start_y + (i as u32) * v_spacing)
            };
            
            node_positions.insert(node.id.clone(), (x, y));
            
            let shape = create_node_shape(node, x, y, node_width, node_height);
            shapes.push(shape);
            _shape_id += 1;
        }
    }
    
    // Create connectors for connections
    for conn in &flowchart.connections {
        if let (Some(&(from_x, from_y)), Some(&(to_x, to_y))) = 
            (node_positions.get(&conn.from), node_positions.get(&conn.to)) 
        {
            let (start_x, start_y, end_x, end_y) = if is_horizontal {
                (from_x + node_width, from_y + node_height / 2,
                 to_x, to_y + node_height / 2)
            } else {
                (from_x + node_width / 2, from_y + node_height,
                 to_x + node_width / 2, to_y)
            };
            
            let connector_type = if (start_x as i32 - end_x as i32).abs() < 100_000 
                                 || (start_y as i32 - end_y as i32).abs() < 100_000 {
                ConnectorType::Straight
            } else {
                ConnectorType::Elbow
            };
            
            let (line_color, line_dash) = match conn.arrow_type {
                ArrowStyle::Thick => ("E65100", LineDash::Solid),
                ArrowStyle::Dotted => ("757575", LineDash::Dash),
                ArrowStyle::Open => ("1565C0", LineDash::Solid),
                ArrowStyle::Arrow => ("1565C0", LineDash::Solid),
            };
            
            let mut connector = Connector::new(connector_type, start_x, start_y, end_x, end_y)
                .with_line(ConnectorLine::new(line_color, 19050).with_dash(line_dash))
                .with_end_arrow(ArrowType::Triangle);
            
            if let Some(label) = &conn.label {
                connector = connector.with_label(label);
            }
            
            connectors.push(connector);
        }
    }
    
    DiagramElements { shapes, connectors }
}

fn get_subgraph_color(index: usize) -> &'static str {
    const COLORS: [&str; 6] = ["E3F2FD", "F3E5F5", "E8F5E9", "FFF3E0", "E0F7FA", "FCE4EC"];
    COLORS[index % COLORS.len()]
}

fn create_node_shape(node: &FlowNode, x: u32, y: u32, width: u32, height: u32) -> Shape {
    let shape_type = match node.shape {
        NodeShape::Rectangle => ShapeType::Rectangle,
        NodeShape::RoundedRect => ShapeType::RoundedRectangle,
        NodeShape::Stadium => ShapeType::RoundedRectangle,
        NodeShape::Diamond => ShapeType::Diamond,
        NodeShape::Circle => ShapeType::Ellipse,
        NodeShape::Hexagon => ShapeType::Hexagon,
    };
    
    let fill_color = match node.shape {
        NodeShape::Diamond => "FFF3E0",
        NodeShape::Circle => "E3F2FD",
        _ => "FFFFFF",
    };
    
    Shape::new(shape_type, x, y, width, height)
        .with_fill(ShapeFill::new(fill_color))
        .with_line(ShapeLine::new("1565C0", 2))
        .with_text(&node.label)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_flowchart_nodes() {
        let code = "flowchart LR\n    A[Start] --> B[Process] --> C[End]";
        let flowchart = parse(code);
        assert_eq!(flowchart.direction, FlowDirection::LeftToRight);
        assert!(!flowchart.nodes.is_empty());
        assert!(!flowchart.connections.is_empty());
    }

    #[test]
    fn test_parse_node_shapes() {
        let (id, node) = parse_node_def("A[Rectangle]");
        assert_eq!(id, "A");
        assert!(node.is_some());
        assert_eq!(node.unwrap().shape, NodeShape::Rectangle);

        let (id, node) = parse_node_def("B(Rounded)");
        assert_eq!(id, "B");
        assert_eq!(node.unwrap().shape, NodeShape::RoundedRect);

        let (id, node) = parse_node_def("C{Diamond}");
        assert_eq!(id, "C");
        assert_eq!(node.unwrap().shape, NodeShape::Diamond);
    }

    #[test]
    fn test_generate_flowchart_shapes() {
        let code = "flowchart LR\n    A[Start] --> B[End]";
        let flowchart = parse(code);
        let elements = generate_elements(&flowchart);
        assert!(!elements.shapes.is_empty());
    }
}
