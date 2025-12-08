//! Shared types for Mermaid diagram parsing

use crate::generator::Shape;
use crate::generator::connectors::Connector;

/// Mermaid diagram types
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum MermaidType {
    Flowchart,
    Sequence,
    Pie,
    Gantt,
    ClassDiagram,
    StateDiagram,
    ErDiagram,
    Mindmap,
    Timeline,
    Journey,
    Quadrant,
    GitGraph,
    Unknown,
}

/// Direction of flowchart layout
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum FlowDirection {
    LeftToRight,  // LR
    RightToLeft,  // RL
    TopToBottom,  // TB/TD
    BottomToTop,  // BT
}

/// A parsed flowchart node
#[derive(Debug, Clone)]
pub struct FlowNode {
    pub id: String,
    pub label: String,
    pub shape: NodeShape,
}

/// Node shape types in Mermaid
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum NodeShape {
    Rectangle,      // [text]
    RoundedRect,    // (text)
    Stadium,        // ([text])
    Diamond,        // {text}
    Circle,         // ((text))
    Hexagon,        // {{text}}
}

/// A connection between nodes
#[derive(Debug, Clone)]
pub struct FlowConnection {
    pub from: String,
    pub to: String,
    pub label: Option<String>,
    pub arrow_type: ArrowStyle,
}

/// Arrow styles
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ArrowStyle {
    Arrow,      // -->
    Open,       // ---
    Dotted,     // -.->
    Thick,      // ==>
}

/// A subgraph grouping
#[derive(Debug, Clone)]
pub struct Subgraph {
    pub name: String,
    pub nodes: Vec<String>,
}

/// Parsed flowchart
#[derive(Debug, Clone)]
pub struct Flowchart {
    pub direction: FlowDirection,
    pub nodes: Vec<FlowNode>,
    pub connections: Vec<FlowConnection>,
    pub subgraphs: Vec<Subgraph>,
}

/// Bounding box for diagram positioning
#[derive(Debug, Clone, Copy)]
pub struct DiagramBounds {
    pub x: u32,
    pub y: u32,
    pub width: u32,
    pub height: u32,
}

impl DiagramBounds {
    pub fn new(x: u32, y: u32, width: u32, height: u32) -> Self {
        Self { x, y, width, height }
    }
    
    /// Calculate bounds from a set of positions and dimensions
    pub fn from_elements(positions: &[(u32, u32, u32, u32)]) -> Option<Self> {
        if positions.is_empty() {
            return None;
        }
        
        let mut min_x = u32::MAX;
        let mut min_y = u32::MAX;
        let mut max_x = 0u32;
        let mut max_y = 0u32;
        
        for &(x, y, w, h) in positions {
            min_x = min_x.min(x);
            min_y = min_y.min(y);
            max_x = max_x.max(x + w);
            max_y = max_y.max(y + h);
        }
        
        Some(Self {
            x: min_x,
            y: min_y,
            width: max_x - min_x,
            height: max_y - min_y,
        })
    }
}

/// Result containing shapes and connectors
pub struct DiagramElements {
    pub shapes: Vec<Shape>,
    pub connectors: Vec<Connector>,
    /// Bounding box of the diagram for positioning
    pub bounds: Option<DiagramBounds>,
    /// Whether elements should be grouped
    pub grouped: bool,
}

impl DiagramElements {
    /// Create from shapes only (calculates bounds automatically)
    pub fn from_shapes(shapes: Vec<Shape>) -> Self {
        let element_bounds: Vec<(u32, u32, u32, u32)> = shapes
            .iter()
            .map(|s| (s.x, s.y, s.width, s.height))
            .collect();
        let bounds = DiagramBounds::from_elements(&element_bounds);
        
        Self {
            shapes,
            connectors: Vec::new(),
            bounds,
            grouped: true,
        }
    }
    
    /// Create from shapes and connectors
    pub fn from_shapes_and_connectors(shapes: Vec<Shape>, connectors: Vec<Connector>) -> Self {
        let element_bounds: Vec<(u32, u32, u32, u32)> = shapes
            .iter()
            .map(|s| (s.x, s.y, s.width, s.height))
            .collect();
        let bounds = DiagramBounds::from_elements(&element_bounds);
        
        Self {
            shapes,
            connectors,
            bounds,
            grouped: true,
        }
    }
    
    /// Create empty diagram elements
    pub fn empty() -> Self {
        Self {
            shapes: Vec::new(),
            connectors: Vec::new(),
            bounds: None,
            grouped: false,
        }
    }
}
