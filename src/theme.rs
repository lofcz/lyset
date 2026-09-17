//! Typography and colour system for the print renderer.
//!
//! All measurements are points or millimetres; colours are hex without `#`.
//! The body font is Calibri: the layout engine substitutes the bundled,
//! metric-compatible Carlito, so PDF pagination matches Word exactly.

use crate::ir::PanelVariant;

pub const FONT_BODY: &str = "Calibri";
pub const FONT_MONO: &str = "Courier New";

pub const PT_BODY: f64 = 10.5;
pub const PT_LEAD: f64 = 11.5;
pub const PT_SMALL: f64 = 9.0;
pub const PT_CAPTION: f64 = 8.5;
pub const PT_LABEL: f64 = 7.5;
pub const PT_KICKER: f64 = 8.0;
pub const PT_H1: f64 = 20.0;
pub const PT_H2: f64 = 13.5;
pub const PT_H3: f64 = 11.0;
pub const PT_META: f64 = 9.0;

pub const LINE_MULTIPLE: f64 = 1.15;
/// Carlito / Calibri natural line height (winAscent + winDescent) in em.
pub const FONT_LINE_EM: f64 = 1.22;

/// Line pitch in points for `size_pt` text at `multiple` × single spacing.
/// Emitted as `atLeast` so Word and the PDF layout engine agree exactly and
/// tall inline content (fractions, images) can still grow the line.
pub fn line_pitch_pt(size_pt: f64, multiple: f64) -> f64 {
    (size_pt * FONT_LINE_EM * multiple * 20.0).round() / 20.0
}

pub const INK: &str = "1F2933";
pub const MUTED: &str = "6B7280";
pub const RULE: &str = "D5DAE0";
pub const SOFT: &str = "F3F5F7";
pub const ANSWER: &str = "1D7A3E";
pub const BLANK_LINE: &str = "6B7280";
pub const DEFAULT_ACCENT: &str = "007A85";

/// A4 portrait.
pub const PAGE_W_MM: f64 = 210.0;
pub const PAGE_H_MM: f64 = 297.0;
/// top, right, bottom, left
pub const MARGINS_MM: [f64; 4] = [16.0, 18.0, 17.0, 18.0];
pub const HEADER_DISTANCE_MM: f64 = 8.0;
pub const FOOTER_DISTANCE_MM: f64 = 8.5;

/// Hanging gutter for list markers.
pub const LIST_INDENT_MM: f64 = 6.0;
/// Hanging gutter for task numbers.
pub const TASK_GUTTER_MM: f64 = 8.5;
/// Hanging gutter for option letters.
pub const OPTION_GUTTER_MM: f64 = 7.0;
/// Width of the accent bar on panels.
pub const PANEL_BAR_MM: f64 = 1.4;

pub struct PanelColors {
    pub bar: &'static str,
    pub fill: &'static str,
}

pub fn panel_colors(variant: PanelVariant) -> PanelColors {
    match variant {
        PanelVariant::Tip | PanelVariant::Note => PanelColors { bar: "3B82F6", fill: "EEF4FD" },
        PanelVariant::Warning => PanelColors { bar: "D97706", fill: "FDF5E6" },
        PanelVariant::Important => PanelColors { bar: "7C3AED", fill: "F4F1FC" },
        PanelVariant::Success => PanelColors { bar: "059669", fill: "EBF8F2" },
        PanelVariant::Story => PanelColors { bar: "B45309", fill: "FBF4E8" },
        PanelVariant::Solution => PanelColors { bar: "1D7A3E", fill: "EEF7F1" },
        PanelVariant::Source => PanelColors { bar: "6B7280", fill: "F3F5F7" },
        PanelVariant::Plain => PanelColors { bar: "9CA3AF", fill: "F5F6F8" },
    }
}

/// Approximate advance of one average Carlito glyph at `pt`, in millimetres.
pub fn char_width_mm(pt: f64) -> f64 {
    pt * 0.5 * 0.3528
}

/// Tasks whose body is at most this many estimated characters (roughly a
/// third of a page) are kept on a single page as a keep-with-next chain.
pub const KEEP_TASK_CHARS: usize = 1100;
