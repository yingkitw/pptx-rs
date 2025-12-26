//! Slide layout types

/// Slide layout types
#[derive(Clone, Debug, Copy, PartialEq, Eq)]
pub enum SlideLayout {
    /// Title only (no content area)
    TitleOnly,
    /// Title and content (bullets)
    TitleAndContent,
    /// Title at top, content fills rest
    TitleAndBigContent,
    /// Blank slide
    Blank,
    /// Title centered on slide
    CenteredTitle,
    /// Two columns: title on left, content on right
    TwoColumn,
}

impl SlideLayout {
    pub fn as_str(&self) -> &'static str {
        match self {
            SlideLayout::TitleOnly => "titleOnly",
            SlideLayout::TitleAndContent => "titleAndContent",
            SlideLayout::TitleAndBigContent => "titleAndBigContent",
            SlideLayout::Blank => "blank",
            SlideLayout::CenteredTitle => "centeredTitle",
            SlideLayout::TwoColumn => "twoColumn",
        }
    }
}

