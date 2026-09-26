//! Print-document IR — the JSON schema this crate renders.
//!
//! Field names, enum values and defaults are the public contract. Unknown
//! fields are rejected so producers and this renderer stay in lockstep.

use serde::Deserialize;

/// `locale` / `kind` are part of the contract but the renderer is currently
/// locale-agnostic (labels arrive pre-localised in the IR).
#[allow(dead_code)]
#[derive(Debug, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct PrintDocument {
    pub version: u32,
    pub locale: String,
    pub kind: DocumentKind,
    pub title: String,
    #[serde(default)]
    pub header: Option<Header>,
    #[serde(default)]
    pub footer: Option<Footer>,
    #[serde(default)]
    pub page: Option<Page>,
    #[serde(default)]
    pub theme: Option<Theme>,
    /// Logo printed at the right end of the footer on every page (page
    /// numbers move left of it). Used for brand / free-tier marks.
    #[serde(default)]
    pub watermark: Option<Watermark>,
    pub blocks: Vec<Block>,
}

#[derive(Debug, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct Watermark {
    pub image: Image,
    /// Printed width; the height follows the image aspect. Default 22 mm.
    #[serde(default, rename = "widthMm")]
    pub width_mm: Option<f64>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum DocumentKind {
    Lesson,
    Worksheet,
    Test,
    Activity,
    Print,
}

#[derive(Debug, Default, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct Header {
    #[serde(default)]
    pub left: Option<String>,
    #[serde(default)]
    pub right: Option<String>,
}

#[derive(Debug, Default, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct Footer {
    #[serde(default)]
    pub left: Option<String>,
    #[serde(default, rename = "pageNumbers")]
    pub page_numbers: Option<bool>,
    #[serde(default, rename = "pageNumberFormat")]
    pub page_number_format: Option<String>,
}

#[derive(Debug, Default, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct Page {
    #[serde(default)]
    pub orientation: Option<Orientation>,
    #[serde(default, rename = "marginsMm")]
    pub margins_mm: Option<[f64; 4]>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum Orientation {
    Portrait,
    Landscape,
}

#[derive(Debug, Default, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct Theme {
    #[serde(default)]
    pub accent: Option<String>,
}

// ---------------------------------------------------------------------------
// Inline
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, Default)]
#[serde(rename_all = "lowercase")]
pub enum Tone {
    #[default]
    Default,
    Muted,
    Accent,
    Answer,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", deny_unknown_fields)]
pub enum Inline {
    Text {
        text: String,
        #[serde(default)]
        bold: Option<bool>,
        #[serde(default)]
        italic: Option<bool>,
        #[serde(default)]
        underline: Option<bool>,
        #[serde(default)]
        strike: Option<bool>,
        #[serde(default)]
        code: Option<bool>,
        #[serde(default)]
        sup: Option<bool>,
        #[serde(default)]
        sub: Option<bool>,
        #[serde(default)]
        tone: Option<Tone>,
        /// Syntax token foreground, six hexadecimal digits (optional #).
        #[serde(default)]
        color: Option<String>,
    },
    /// Inline formula. `tex` is the authored LaTeX (kept for text fallbacks
    /// and diagnostics); `mathml` is optional presentation MathML, converted
    /// to OMML when present.
    Math {
        tex: String,
        #[serde(default)]
        mathml: Option<String>,
    },
    Break,
    Blank {
        width: u32,
        #[serde(default)]
        answer: Option<String>,
        #[serde(default)]
        reveal: Option<bool>,
    },
    Link {
        href: String,
        text: String,
    },
}

// ---------------------------------------------------------------------------
// Blocks
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum Align {
    Left,
    Center,
    Right,
    Justify,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, Default)]
#[serde(rename_all = "lowercase")]
pub enum ParagraphStyle {
    #[default]
    Body,
    Lead,
    Small,
    Caption,
    Label,
    Code,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum PanelVariant {
    Note,
    Tip,
    Warning,
    Important,
    Success,
    Story,
    Solution,
    Source,
    Plain,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, Default)]
#[serde(rename_all = "lowercase")]
pub enum TableStyle {
    Grid,
    #[default]
    Rules,
    Plain,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, Default)]
#[serde(rename_all = "lowercase")]
pub enum VAlign {
    #[default]
    Top,
    Middle,
    Bottom,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, Default)]
#[serde(rename_all = "lowercase")]
pub enum RuleTone {
    #[default]
    Muted,
    Accent,
}

/// Picture bytes come either inline (`data`, base64) or from a file the
/// caller controls (`path`) — exactly one of the two. Large exports use
/// `path` so the IR stays small.
///
/// `alt` is accepted for schema parity; DOCX pictures carry no alt text yet.
#[allow(dead_code)]
#[derive(Debug, Clone, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct Image {
    #[serde(default)]
    pub data: Option<String>,
    #[serde(default)]
    pub path: Option<std::path::PathBuf>,
    pub mime: String,
    #[serde(default)]
    pub alt: Option<String>,
    #[serde(default)]
    pub caption: Option<String>,
    #[serde(default, rename = "maxWidth")]
    pub max_width: Option<f64>,
    #[serde(default, rename = "maxHeightMm")]
    pub max_height_mm: Option<f64>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct ListItem {
    pub content: Vec<Inline>,
    #[serde(default)]
    pub number: Option<i64>,
    #[serde(default)]
    pub children: Option<Vec<Block>>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct TableCell {
    pub blocks: Vec<Block>,
    #[serde(default, rename = "colSpan")]
    pub col_span: Option<u32>,
    #[serde(default)]
    pub align: Option<Align>,
    #[serde(default)]
    pub shade: Option<bool>,
    #[serde(default, rename = "vAlign")]
    pub v_align: Option<VAlign>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct FactRow {
    pub label: String,
    pub blocks: Vec<Block>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct Field {
    pub label: String,
    #[serde(rename = "widthMm")]
    pub width_mm: f64,
    #[serde(default)]
    pub value: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct OptionItem {
    pub label: String,
    pub content: Vec<Inline>,
    #[serde(default)]
    pub correct: Option<bool>,
    #[serde(default)]
    pub image: Option<Image>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct CardPart {
    pub blocks: Vec<Block>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct CardItem {
    pub parts: Vec<CardPart>,
    #[serde(default)]
    pub badge: Option<String>,
}

#[derive(Debug, Clone, Default, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct GridCell {
    #[serde(default)]
    pub char: Option<String>,
    #[serde(default)]
    pub mark: Option<bool>,
    #[serde(default)]
    pub blocked: Option<bool>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum GridAlign {
    Left,
    Center,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", deny_unknown_fields)]
pub enum Block {
    Heading {
        level: u8,
        content: Vec<Inline>,
        #[serde(default)]
        kicker: Option<String>,
        #[serde(default)]
        meta: Option<String>,
        #[serde(default, rename = "pageBreakBefore")]
        page_break_before: Option<bool>,
    },
    Paragraph {
        content: Vec<Inline>,
        #[serde(default)]
        style: Option<ParagraphStyle>,
        #[serde(default)]
        align: Option<Align>,
        #[serde(default)]
        trailing: Option<Vec<Inline>>,
        #[serde(default)]
        indent: Option<u32>,
        #[serde(default, rename = "keepWithNext")]
        keep_with_next: Option<bool>,
    },
    List {
        ordered: bool,
        items: Vec<ListItem>,
        #[serde(default)]
        tight: Option<bool>,
    },
    Table {
        rows: Vec<Vec<TableCell>>,
        #[serde(default)]
        columns: Option<Vec<f64>>,
        #[serde(default, rename = "headerRows")]
        header_rows: Option<u32>,
        #[serde(default)]
        style: Option<TableStyle>,
        #[serde(default)]
        dense: Option<bool>,
    },
    Panel {
        variant: PanelVariant,
        #[serde(default)]
        label: Option<String>,
        blocks: Vec<Block>,
    },
    Image {
        image: Image,
        #[serde(default)]
        align: Option<Align>,
    },
    Gallery {
        images: Vec<Image>,
    },
    Facts {
        rows: Vec<FactRow>,
    },
    Rule {
        #[serde(default)]
        tone: Option<RuleTone>,
    },
    Spacer {
        #[serde(rename = "heightMm")]
        height_mm: f64,
    },
    PageBreak,
    AnswerLines {
        lines: u32,
    },
    AnswerBox {
        #[serde(rename = "heightMm")]
        height_mm: f64,
    },
    Fields {
        items: Vec<Field>,
    },
    Task {
        #[serde(default)]
        number: Option<i64>,
        #[serde(default)]
        points: Option<String>,
        prompt: Vec<Inline>,
        blocks: Vec<Block>,
    },
    Options {
        #[serde(default)]
        columns: Option<u8>,
        items: Vec<OptionItem>,
        #[serde(default)]
        reveal: Option<bool>,
    },
    Cards {
        columns: u8,
        #[serde(default, rename = "heightMm")]
        height_mm: Option<f64>,
        #[serde(default)]
        cut: Option<bool>,
        items: Vec<CardItem>,
    },
    Grid {
        rows: Vec<Vec<Option<GridCell>>>,
        #[serde(default, rename = "rowLabels")]
        row_labels: Option<Vec<String>>,
        #[serde(default, rename = "cellMm")]
        cell_mm: Option<f64>,
        #[serde(default)]
        align: Option<GridAlign>,
    },
}
