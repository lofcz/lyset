//! IR → rdocx `Document`.
//!
//! The renderer owns every visual decision (type scale, spacing, colours,
//! gutters). Content flows into either the document body or a table cell
//! (panels, facts, galleries, option grids are all tables), abstracted by
//! [`Sink`]. Widths are tracked explicitly in millimetres because Word tables
//! need fixed column widths to paginate identically in Word and in the PDF
//! layout engine.

use std::collections::HashMap;
use std::hash::{Hash, Hasher};

use base64::Engine;
use rdocx::{
    Alignment, BorderStyle, Cell, Document, HdrFtrType, Length, OfficeMath, Paragraph, Table,
    TabAlignment, UnderlineStyle, VerticalAlignment, equation_from_latex, equation_from_mathml,
};

use crate::ir::{
    Align, Block, CardItem, Field, GridAlign, GridCell, Image, Inline, ListItem, OptionItem,
    Orientation, PanelVariant, ParagraphStyle, PrintDocument, TableCell, TableStyle, Tone, VAlign,
};
use crate::theme::{self as t, panel_colors};

// ---------------------------------------------------------------------------
// Public entry
// ---------------------------------------------------------------------------

pub struct RenderReport {
    pub warnings: Vec<String>,
}

pub fn render(ir: &PrintDocument) -> Result<(Document, RenderReport), String> {
    if ir.version != 1 {
        return Err(format!("unsupported IR version {}", ir.version));
    }
    let mut doc = Document::new();
    let mut ctx = Ctx::new(ir)?;

    // Page geometry.
    let (page_w, page_h) = match ir.page.as_ref().and_then(|p| p.orientation) {
        Some(Orientation::Landscape) => (t::PAGE_H_MM, t::PAGE_W_MM),
        _ => (t::PAGE_W_MM, t::PAGE_H_MM),
    };
    let m = ir.page.as_ref().and_then(|p| p.margins_mm).unwrap_or(t::MARGINS_MM);
    doc.set_page_size(Length::mm(page_w), Length::mm(page_h));
    doc.set_margins(Length::mm(m[0]), Length::mm(m[1]), Length::mm(m[2]), Length::mm(m[3]));
    doc.set_header_footer_distance(Length::mm(t::HEADER_DISTANCE_MM), Length::mm(t::FOOTER_DISTANCE_MM));
    ctx.content_w = page_w - m[1] - m[3];
    doc.set_title(&ir.title);
    doc.set_author("lyset");

    // Pre-embed every image and hyperlink so cells can reference them.
    collect_assets(&mut doc, &mut ctx, &ir.blocks);

    let frame = Frame::new(ctx.content_w);
    render_blocks(&mut doc, &mut ctx, &frame, &ir.blocks);

    install_header_footer(&mut doc, &mut ctx, ir);

    if !ctx.math_errors.is_empty() {
        return Err(format!("cannot export document with invalid or lossy math: {}", ctx.math_errors.join("; ")));
    }
    Ok((doc, RenderReport { warnings: ctx.warnings }))
}

// ---------------------------------------------------------------------------
// Context
// ---------------------------------------------------------------------------

struct EmbeddedImage {
    bytes: Vec<u8>,
    filename: String,
    rel_id: String,
    width_px: u32,
    height_px: u32,
}

struct Ctx {
    accent: String,
    content_w: f64,
    images: HashMap<u64, EmbeddedImage>,
    links: HashMap<String, String>,
    warnings: Vec<String>,
    math_errors: Vec<String>,
    image_seq: usize,
    /// Keep-with-next policy for paragraphs created by the block being
    /// rendered (small tasks stay on one page as a chain).
    keep: Keep,
}

/// Which of a block's paragraphs / table rows get `keepNext`.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Keep {
    Off,
    /// Every paragraph and row: the block is followed by more of the same task.
    All,
    /// All but the final paragraph / row: the block closes the chain.
    ExceptLast,
}

impl Ctx {
    fn new(ir: &PrintDocument) -> Result<Self, String> {
        let accent = ir.theme.as_ref().and_then(|th| th.accent.as_deref())
            .unwrap_or(t::DEFAULT_ACCENT).trim();
        let accent = accent.strip_prefix('#').unwrap_or(accent);
        if !matches!(accent.len(), 3 | 6) || !accent.bytes().all(|b| b.is_ascii_hexdigit()) {
            return Err("theme accent must be a 3- or 6-digit hexadecimal colour".to_string());
        }
        // OOXML requires six hex digits, even when the input uses CSS shorthand.
        let accent = if accent.len() == 3 {
            accent.chars().flat_map(|c| [c, c]).collect::<String>()
        } else {
            accent.to_string()
        }.to_uppercase();
        Ok(Ctx {
            accent,
            content_w: 0.0,
            images: HashMap::new(),
            links: HashMap::new(),
            warnings: Vec::new(),
            math_errors: Vec::new(),
            image_seq: 0,
            keep: Keep::Off,
        })
    }

    fn tone_color(&self, tone: Tone) -> String {
        match tone {
            Tone::Default => t::INK.to_string(),
            Tone::Muted => t::MUTED.to_string(),
            Tone::Accent => self.accent.clone(),
            Tone::Answer => t::ANSWER.to_string(),
        }
    }
}

/// Horizontal extent available to the content being rendered.
#[derive(Clone, Copy)]
struct Frame {
    /// Full width of the container in mm (body or cell).
    width: f64,
    /// Left indent inside that container (tasks, nested lists).
    indent: f64,
    /// Nesting depth of lists (drives marker style: 1. → a. → i.).
    list_depth: u8,
}

impl Frame {
    fn new(width: f64) -> Frame {
        Frame { width, indent: 0.0, list_depth: 0 }
    }
    fn inner_width(&self) -> f64 {
        (self.width - self.indent).max(10.0)
    }
    fn indented(&self, by: f64) -> Frame {
        Frame { width: self.width, indent: self.indent + by, list_depth: self.list_depth }
    }
    fn nested_list(&self, by: f64) -> Frame {
        Frame { width: self.width, indent: self.indent + by, list_depth: self.list_depth.saturating_add(1) }
    }
}

fn ordered_marker(depth: u8, n: i64) -> String {
    match depth % 3 {
        0 => format!("{n}."),
        1 => {
            let idx = ((n - 1).max(0) % 26) as u8;
            format!("{}.", (b'a' + idx) as char)
        }
        _ => format!("{}.", roman_lower(n)),
    }
}

fn roman_lower(mut n: i64) -> String {
    if n <= 0 || n >= 4000 {
        return n.to_string();
    }
    const T: [(i64, &str); 13] = [
        (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i"),
    ];
    let mut out = String::new();
    for (v, r) in T {
        while n >= v {
            out.push_str(r);
            n -= v;
        }
    }
    out
}

fn image_key(image: &Image) -> u64 {
    let mut h = std::collections::hash_map::DefaultHasher::new();
    image.data.hash(&mut h);
    image.path.hash(&mut h);
    h.finish()
}

/// Raw picture bytes from `data` (base64) or `path`; the message names the
/// source so a warning points at the right thing.
fn image_bytes(image: &Image) -> Result<Vec<u8>, String> {
    match (&image.data, &image.path) {
        (Some(data), None) => base64::engine::general_purpose::STANDARD
            .decode(data.trim())
            .map_err(|e| format!("invalid base64 ({e})")),
        (None, Some(path)) => std::fs::read(path).map_err(|e| format!("read {}: {e}", path.display())),
        (Some(_), Some(_)) => Err("both data and path given".to_string()),
        (None, None) => Err("neither data nor path given".to_string()),
    }
}

fn collect_assets(doc: &mut Document, ctx: &mut Ctx, blocks: &[Block]) {
    for block in blocks {
        match block {
            Block::Image { image, .. } => embed_image(doc, ctx, image),
            Block::Gallery { images } => images.iter().for_each(|i| embed_image(doc, ctx, i)),
            Block::Options { items, .. } => {
                for item in items {
                    collect_inline_assets(doc, ctx, &item.content);
                    if let Some(img) = &item.image {
                        embed_image(doc, ctx, img);
                    }
                }
            }
            Block::Heading { content, .. } => collect_inline_assets(doc, ctx, content),
            Block::Paragraph { content, trailing, .. } => {
                collect_inline_assets(doc, ctx, content);
                if let Some(tr) = trailing {
                    collect_inline_assets(doc, ctx, tr);
                }
            }
            Block::List { items, .. } => collect_list_assets(doc, ctx, items),
            Block::Table { rows, .. } => {
                for row in rows {
                    for cell in row {
                        collect_assets(doc, ctx, &cell.blocks);
                    }
                }
            }
            Block::Cards { items, .. } => {
                for item in items {
                    for part in &item.parts {
                        collect_assets(doc, ctx, &part.blocks);
                    }
                }
            }
            Block::Panel { blocks, .. } => collect_assets(doc, ctx, blocks),
            Block::Facts { rows } => rows.iter().for_each(|r| collect_assets(doc, ctx, &r.blocks)),
            Block::Task { prompt, blocks, .. } => {
                collect_inline_assets(doc, ctx, prompt);
                collect_assets(doc, ctx, blocks);
            }
            _ => {}
        }
    }
}

fn collect_list_assets(doc: &mut Document, ctx: &mut Ctx, items: &[ListItem]) {
    for item in items {
        collect_inline_assets(doc, ctx, &item.content);
        if let Some(children) = &item.children {
            collect_assets(doc, ctx, children);
        }
    }
}

fn collect_inline_assets(doc: &mut Document, ctx: &mut Ctx, inlines: &[Inline]) {
    for inline in inlines {
        if let Inline::Link { href, .. } = inline {
            if !ctx.links.contains_key(href) {
                let rel = doc.add_hyperlink_relationship(href);
                ctx.links.insert(href.clone(), rel);
            }
        }
    }
}

fn embed_image(doc: &mut Document, ctx: &mut Ctx, image: &Image) {
    let key = image_key(image);
    if ctx.images.contains_key(&key) {
        return;
    }
    let bytes = match image_bytes(image) {
        Ok(b) => b,
        Err(e) => {
            ctx.warnings.push(format!("image: {e}"));
            return;
        }
    };
    let Some(info) = oxml_media::probe(&bytes) else {
        ctx.warnings.push(format!("image: unsupported or corrupt data ({} bytes, {})", bytes.len(), image.mime));
        return;
    };
    let ext = match image.mime.as_str() {
        "image/jpeg" => "jpg",
        "image/gif" => "gif",
        "image/webp" => "webp",
        "image/svg+xml" => "svg",
        _ => "png",
    };
    ctx.image_seq += 1;
    let filename = format!("figure-{}.{ext}", ctx.image_seq);
    let rel_id = doc.embed_image(&bytes, &filename);
    ctx.images.insert(
        key,
        EmbeddedImage { bytes, filename, rel_id, width_px: info.width_px, height_px: info.height_px },
    );
}

// ---------------------------------------------------------------------------
// Sink: body or cell
// ---------------------------------------------------------------------------

trait Sink {
    fn para(&mut self) -> Paragraph<'_>;
    fn table(&mut self, rows: usize, cols: usize) -> Table<'_>;
    /// Adds a picture paragraph; returns it for alignment / spacing.
    fn picture(&mut self, img: &EmbeddedImage, w: Length, h: Length) -> Paragraph<'_>;
    /// Direct paragraphs of this container (not those inside nested tables).
    fn direct_paragraph_count(&self) -> usize;
    fn direct_paragraph(&mut self, index: usize) -> Option<Paragraph<'_>>;
}

impl Sink for Document {
    fn para(&mut self) -> Paragraph<'_> {
        self.add_paragraph("")
    }
    fn table(&mut self, rows: usize, cols: usize) -> Table<'_> {
        self.add_table(rows, cols)
    }
    fn picture(&mut self, img: &EmbeddedImage, w: Length, h: Length) -> Paragraph<'_> {
        self.add_picture(&img.bytes, &img.filename, w, h)
    }
    fn direct_paragraph_count(&self) -> usize {
        self.paragraph_count()
    }
    fn direct_paragraph(&mut self, index: usize) -> Option<Paragraph<'_>> {
        self.paragraph_mut(index)
    }
}

impl Sink for Cell<'_> {
    fn para(&mut self) -> Paragraph<'_> {
        self.add_paragraph("")
    }
    fn table(&mut self, rows: usize, cols: usize) -> Table<'_> {
        self.add_table(rows, cols)
    }
    fn picture(&mut self, img: &EmbeddedImage, w: Length, h: Length) -> Paragraph<'_> {
        self.add_picture(&img.rel_id, w, h);
        let last = self.paragraph_count().saturating_sub(1);
        self.paragraph_mut(last).expect("add_picture appends a paragraph")
    }
    fn direct_paragraph_count(&self) -> usize {
        self.paragraph_count()
    }
    fn direct_paragraph(&mut self, index: usize) -> Option<Paragraph<'_>> {
        self.paragraph_mut(index)
    }
}

/// Apply the current keep policy to a finished table: Word keeps a row with
/// the next one when every paragraph in the row has `keepNext`.
fn keep_table_rows(table: &mut Table<'_>, rows: usize, cols: usize, keep: Keep) {
    let limit = match keep {
        Keep::Off => return,
        Keep::All => rows,
        Keep::ExceptLast => rows.saturating_sub(1),
    };
    for r in 0..limit {
        for c in 0..cols {
            let Some(mut cell) = table.cell(r, c) else { continue };
            for i in 0..cell.paragraph_count() {
                if let Some(mut p) = cell.paragraph_mut(i) {
                    p.set_keep_with_next(true);
                }
            }
        }
    }
}

// ---------------------------------------------------------------------------
// Run styling
// ---------------------------------------------------------------------------

#[derive(Clone)]
struct RunStyle {
    size: f64,
    bold: bool,
    italic: bool,
    color: String,
}

impl RunStyle {
    fn body() -> Self {
        RunStyle { size: t::PT_BODY, bold: false, italic: false, color: t::INK.to_string() }
    }
    fn with_size(mut self, size: f64) -> Self {
        self.size = size;
        self
    }
    fn bold(mut self) -> Self {
        self.bold = true;
        self
    }
    fn italic(mut self) -> Self {
        self.italic = true;
        self
    }
    fn color(mut self, hex: &str) -> Self {
        self.color = hex.to_string();
        self
    }
}

fn styled_run<'p>(p: &'p mut Paragraph<'_>, text: &str, style: &RunStyle) -> rdocx::Run<'p> {
    let mut run = p.add_run(text);
    run.set_font(t::FONT_BODY);
    run.set_size(style.size);
    run.set_color(&style.color);
    if style.bold {
        run.set_bold(true);
    }
    if style.italic {
        run.set_italic(true);
    }
    run
}

fn write_inlines(ctx: &mut Ctx, p: &mut Paragraph<'_>, inlines: &[Inline], base: &RunStyle) {
    write_inlines_inner(ctx, p, inlines, base);
    // The layout engine sizes an equation from the run that follows it (or
    // the paragraph mark, which carries no size here). A trailing empty run
    // in the base style keeps math-only cells and math-final paragraphs at
    // the paragraph's text size.
    if matches!(inlines.last(), Some(Inline::Math { .. })) {
        styled_run(p, "", base);
    }
}

fn write_inlines_inner(ctx: &mut Ctx, p: &mut Paragraph<'_>, inlines: &[Inline], base: &RunStyle) {
    for inline in inlines {
        match inline {
            Inline::Text { text, bold, italic, underline, strike, code, sup, sub, tone } => {
                let mut style = base.clone();
                if bold.unwrap_or(false) {
                    style.bold = true;
                }
                if italic.unwrap_or(false) {
                    style.italic = true;
                }
                if let Some(tone) = tone {
                    if *tone != Tone::Default {
                        style.color = ctx.tone_color(*tone);
                    }
                }
                let mut run = styled_run(p, text, &style);
                if code.unwrap_or(false) {
                    run.set_font(t::FONT_MONO);
                    run.set_size(style.size - 1.0);
                }
                if underline.unwrap_or(false) {
                    run.set_underline(true);
                }
                if strike.unwrap_or(false) {
                    run.set_strike(true);
                }
                if sup.unwrap_or(false) {
                    run.set_superscript();
                }
                if sub.unwrap_or(false) {
                    run.set_subscript();
                }
            }
            Inline::Math { tex, mathml } => write_math(ctx, p, tex, mathml.as_deref()),
            Inline::Break => p.add_line_break(),
            Inline::Blank { width, answer, reveal } => {
                let revealed = reveal.unwrap_or(false);
                match (revealed, answer.as_deref()) {
                    (true, Some(ans)) if !ans.is_empty() => {
                        let style = base.clone().bold().color(t::ANSWER);
                        let mut run = styled_run(p, ans, &style);
                        run.set_underline_style(UnderlineStyle::Single);
                    }
                    _ => {
                        let line = "_".repeat((*width).max(3) as usize);
                        let style = base.clone().color(t::BLANK_LINE);
                        styled_run(p, &line, &style);
                    }
                }
            }
            Inline::Link { href, text } => {
                let rel = ctx.links.get(href).cloned();
                let label = if text.is_empty() { href } else { text };
                match rel {
                    Some(rel_id) => {
                        let mut run = p.add_hyperlink(label, &rel_id);
                        run.set_font(t::FONT_BODY);
                        run.set_size(base.size);
                        run.set_color(&ctx.accent);
                        run.set_underline(true);
                    }
                    None => {
                        let style = base.clone().color(&ctx.accent);
                        styled_run(p, label, &style);
                    }
                }
            }
        }
    }
}

/// Insert an editable Office Math object. Any lossy conversion prevents export.
fn write_math(ctx: &mut Ctx, p: &mut Paragraph<'_>, tex: &str, mathml: Option<&str>) {
    let converted = match mathml {
        Some(mathml) => equation_from_mathml(mathml),
        None => equation_from_latex(tex),
    };
    match converted {
        Ok(result) => {
            if !result.diagnostics.is_empty() {
                for d in &result.diagnostics {
                    ctx.math_errors.push(format!("`{tex}` at {}: {}", d.path, d.message));
                }
                return;
            }
            let om = OfficeMath::inline(result.value.expressions);
            if let Err(e) = p.add_equation(om) {
                ctx.math_errors.push(format!("could not insert `{tex}`: {e}"));
            }
        }
        Err(e) => ctx.math_errors.push(format!("unsupported formula `{tex}`: {e}")),
    }
}

/// Real `<w:tab/>` (a literal tab inside `<w:t>` renders as a missing glyph).
fn tab(p: &mut Paragraph<'_>, style: &RunStyle) {
    let mut run = styled_run(p, "", style);
    run.add_tab();
}

fn set_align(p: &mut Paragraph<'_>, align: Option<Align>) {
    if let Some(a) = align {
        p.set_alignment(match a {
            Align::Left => Alignment::Left,
            Align::Center => Alignment::Center,
            Align::Right => Alignment::Right,
            Align::Justify => Alignment::Justify,
        });
    }
}

fn spacing(p: &mut Paragraph<'_>, before_pt: f64, after_pt: f64) {
    spacing_sized(p, before_pt, after_pt, t::PT_BODY, t::LINE_MULTIPLE);
}

fn spacing_sized(p: &mut Paragraph<'_>, before_pt: f64, after_pt: f64, size_pt: f64, multiple: f64) {
    p.set_space_before(Length::pt(before_pt));
    p.set_space_after(Length::pt(after_pt));
    p.set_line_spacing_at_least(t::line_pitch_pt(size_pt, multiple));
}

/// A near-invisible paragraph used to separate tables from following text.
fn gap<S: Sink>(sink: &mut S, pt: f64) {
    let mut p = sink.para();
    p.set_space_before(Length::pt(0.0));
    p.set_space_after(Length::pt(0.0));
    p.set_line_spacing(pt);
    let mut run = p.add_run("");
    run.set_size(pt.max(1.0));
    run.set_font(t::FONT_BODY);
}

// ---------------------------------------------------------------------------
// Blocks
// ---------------------------------------------------------------------------

fn render_blocks<S: Sink>(sink: &mut S, ctx: &mut Ctx, frame: &Frame, blocks: &[Block]) {
    for block in blocks {
        render_block(sink, ctx, frame, block);
    }
}

fn render_block<S: Sink>(sink: &mut S, ctx: &mut Ctx, frame: &Frame, block: &Block) {
    match block {
        Block::Heading { level, content, kicker, meta, page_break_before } => {
            render_heading(sink, ctx, frame, *level, content, kicker.as_deref(), meta.as_deref(), page_break_before.unwrap_or(false));
        }
        Block::Paragraph { content, style, align, trailing, indent, keep_with_next } => {
            let style = style.unwrap_or_default();
            let extra_indent = indent.unwrap_or(0) as f64 * t::LIST_INDENT_MM;
            let f = frame.indented(extra_indent);
            let mut p = sink.para();
            apply_paragraph_style(&mut p, &f, style);
            set_align(&mut p, *align);
            if keep_with_next.unwrap_or(false) {
                p.set_keep_with_next(true);
            }
            let run_style = run_style_for(style);
            if style == ParagraphStyle::Label {
                let upper: Vec<Inline> = content.iter().map(uppercase_inline).collect();
                write_inlines(ctx, &mut p, &upper, &run_style);
            } else {
                write_inlines(ctx, &mut p, content, &run_style);
            }
            if let Some(tr) = trailing {
                if !tr.is_empty() {
                    p.set_add_tab_stop(TabAlignment::Right, Length::mm(f.width));
                    tab(&mut p, &run_style);
                    write_inlines(ctx, &mut p, tr, &run_style.clone().with_size(t::PT_META));
                }
            }
        }
        Block::List { ordered, items, tight } => {
            render_list(sink, ctx, frame, *ordered, items, tight.unwrap_or(false));
        }
        Block::Table { rows, columns, header_rows, style, dense } => {
            render_table(sink, ctx, frame, rows, columns.as_deref(), header_rows.unwrap_or(0) as usize, style.unwrap_or_default(), dense.unwrap_or(false));
        }
        Block::Panel { variant, label, blocks } => render_panel(sink, ctx, frame, *variant, label.as_deref(), blocks),
        Block::Image { image, align } => render_image(sink, ctx, frame, image, *align, None),
        Block::Gallery { images } => render_gallery(sink, ctx, frame, images),
        Block::Facts { rows } => render_facts(sink, ctx, frame, rows),
        Block::Rule { tone } => {
            let accent = matches!(tone, Some(crate::ir::RuleTone::Accent));
            let mut p = sink.para();
            p.set_space_before(Length::pt(2.0));
            p.set_space_after(Length::pt(6.0));
            p.set_line_spacing(2.0);
            if frame.indent > 0.0 {
                p.set_indent_left(Length::mm(frame.indent));
            }
            let (sz, color) = if accent { (12, ctx.accent.clone()) } else { (6, t::RULE.to_string()) };
            p.set_border_bottom(BorderStyle::Single, sz, &color);
            let mut r = p.add_run("");
            r.set_size(2.0);
        }
        Block::Spacer { height_mm } => {
            let mut p = sink.para();
            p.set_space_before(Length::pt(0.0));
            p.set_space_after(Length::pt(0.0));
            p.set_line_spacing(height_mm * 72.0 / 25.4);
            let mut r = p.add_run("");
            r.set_size(2.0);
        }
        Block::PageBreak => {
            let mut p = sink.para();
            p.set_page_break_before(true);
            p.set_space_before(Length::pt(0.0));
            p.set_space_after(Length::pt(0.0));
            p.set_line_spacing(1.0);
            let mut r = p.add_run("");
            r.set_size(1.0);
        }
        Block::AnswerLines { lines } => {
            for i in 0..*lines {
                let mut p = sink.para();
                p.set_space_before(Length::pt(if i == 0 { 2.0 } else { 0.0 }));
                p.set_space_after(Length::pt(0.0));
                p.set_line_spacing(8.0 * 72.0 / 25.4);
                if frame.indent > 0.0 {
                    p.set_indent_left(Length::mm(frame.indent));
                }
                p.set_border_bottom(BorderStyle::Dotted, 6, t::BLANK_LINE);
                let mut r = p.add_run("");
                r.set_size(t::PT_BODY);
                r.set_font(t::FONT_BODY);
            }
            gap(sink, 3.0);
        }
        Block::AnswerBox { height_mm } => {
            let w = frame.inner_width();
            let mut table = sink.table(1, 1);
            table.set_width(Length::mm(w));
            table.set_layout_fixed();
            table.set_column_width(0, Length::mm(w));
            if frame.indent > 0.0 {
                table.set_indent(Length::mm(frame.indent));
            }
            table.set_borders(BorderStyle::Single, 6, t::BLANK_LINE);
            if let Some(mut row) = table.row(0) {
                row.set_height_exact(Length::mm(*height_mm));
            }
            gap(sink, 4.0);
        }
        Block::Fields { items } => render_fields(sink, frame, items),
        Block::Task { number, points, prompt, blocks } => render_task(sink, ctx, frame, *number, points.as_deref(), prompt, blocks),
        Block::Options { columns, items, reveal } => render_options(sink, ctx, frame, *columns, items, reveal.unwrap_or(false)),
        Block::Cards { columns, height_mm, cut, items } => render_cards(sink, ctx, frame, *columns, *height_mm, cut.unwrap_or(false), items),
        Block::Grid { rows, row_labels, cell_mm, align } => render_grid(sink, ctx, frame, rows, row_labels.as_deref(), *cell_mm, *align),
    }
}

// ---------------------------------------------------------------------------
// Cut-out cards
// ---------------------------------------------------------------------------

/// Card sheets: `columns` cards per table row, every card in a row sharing a
/// height so the sheet cuts cleanly. Parts of one card stack as table rows so
/// the border between them is the cut line between the halves. The table has
/// no borders of its own — each card cell frames itself — so trailing empty
/// cells of a short final row stay blank paper.
fn render_cards<S: Sink>(sink: &mut S, ctx: &mut Ctx, frame: &Frame, columns: u8, height_mm: Option<f64>, cut: bool, items: &[CardItem]) {
    if items.is_empty() {
        return;
    }
    let cols = (columns.max(1) as usize).min(4);
    let total = frame.inner_width();
    let gutter = if cut { 0.0 } else { 3.0 };
    let col_w = (total - gutter * (cols as f64 - 1.0)) / cols as f64;
    let pad_h = 2.6;
    let pad_v = 2.0;
    let (style, sz, color) = if cut { (BorderStyle::Dashed, 6, t::MUTED) } else { (BorderStyle::Single, 8, ctx.accent.as_str()) };
    let color = color.to_string();

    for chunk in items.chunks(cols) {
        let parts = chunk.iter().map(|c| c.parts.len().max(1)).max().unwrap_or(1);
        // Solid frames get a gutter column between cards; cut sheets share edges.
        let grid_cols = if cut { cols } else { cols * 2 - 1 };
        let mut table = sink.table(parts, grid_cols);
        table.set_width(Length::mm(total));
        table.set_layout_fixed();
        if frame.indent > 0.0 {
            table.set_indent(Length::mm(frame.indent));
        }
        for g in 0..grid_cols {
            let w = if !cut && g % 2 == 1 { gutter } else { col_w };
            table.set_column_width(g, Length::mm(w));
        }
        table.set_borders(BorderStyle::None, 0, "auto");
        table.set_cell_margins(Length::mm(pad_v), Length::mm(pad_h), Length::mm(pad_v), Length::mm(pad_h));
        let part_h = height_mm.map(|h| h / parts as f64);
        for r in 0..parts {
            if let Some(mut row) = table.row(r) {
                row.set_cant_split();
                if let Some(h) = part_h {
                    row.set_height(Length::mm(h));
                }
            }
        }
        for (ci, card) in chunk.iter().enumerate() {
            let g = if cut { ci } else { ci * 2 };
            for r in 0..parts {
                let Some(mut cell) = table.cell(r, g) else { continue };
                cell.set_width(Length::mm(col_w));
                cell.set_borders(style, sz, &color);
                cell.set_vertical_alignment(VerticalAlignment::Center);
                let inner = Frame::new(col_w - 2.0 * pad_h);
                let before = cell.paragraph_count();
                if r == 0 {
                    if let Some(badge) = card.badge.as_deref().filter(|b| !b.trim().is_empty()) {
                        let mut p = cell.para();
                        spacing(&mut p, 0.0, 1.0);
                        p.set_alignment(Alignment::Right);
                        styled_run(&mut p, badge, &RunStyle::body().with_size(t::PT_LABEL).color(t::MUTED));
                    }
                }
                if let Some(part) = card.parts.get(r) {
                    render_cell_blocks(&mut cell, ctx, &inner, &part.blocks, false, Some(Align::Center));
                }
                if cell.paragraph_count() > before {
                    cell.remove_first_empty_paragraph();
                }
            }
        }
        // A page break between two card rows is fine; between the halves of
        // one card it is not, so chain the part rows with keep-with-next.
        if parts > 1 {
            keep_table_rows(&mut table, parts, grid_cols, Keep::ExceptLast);
        }
        gap(sink, if cut { 0.0 } else { 3.0 });
    }
    gap(sink, 4.0);
}

// ---------------------------------------------------------------------------
// Letter grid
// ---------------------------------------------------------------------------

/// Square letter grid. Squares frame themselves (table borders off) so `null`
/// gaps stay blank paper; the optional label column sits left of the grid.
fn render_grid<S: Sink>(sink: &mut S, ctx: &mut Ctx, frame: &Frame, rows: &[Vec<Option<GridCell>>], row_labels: Option<&[String]>, cell_mm: Option<f64>, align: Option<GridAlign>) {
    let width = rows.iter().map(|r| r.len()).max().unwrap_or(0);
    if width == 0 || rows.is_empty() {
        return;
    }
    let total = frame.inner_width();
    let has_labels = row_labels.map(|l| l.iter().any(|s| !s.trim().is_empty())).unwrap_or(false);
    let label_w = if has_labels { 9.0 } else { 0.0 };
    let fit = (total - label_w) / width as f64;
    let cell = cell_mm.unwrap_or_else(|| fit.clamp(5.0, 9.0)).min(fit);
    let grid_w = label_w + cell * width as f64;
    let indent = match align.unwrap_or(GridAlign::Left) {
        GridAlign::Left => frame.indent,
        GridAlign::Center => frame.indent + ((total - grid_w) / 2.0).max(0.0),
    };
    let cols = width + usize::from(has_labels);
    let letter_pt = (cell * 72.0 / 25.4 * 0.55).clamp(7.0, 14.0);

    let mut table = sink.table(rows.len(), cols);
    table.set_width(Length::mm(grid_w));
    table.set_layout_fixed();
    if indent > 0.0 {
        table.set_indent(Length::mm(indent));
    }
    if has_labels {
        table.set_column_width(0, Length::mm(label_w));
    }
    for c in 0..width {
        table.set_column_width(c + usize::from(has_labels), Length::mm(cell));
    }
    table.set_borders(BorderStyle::None, 0, "auto");
    table.set_cell_margins(Length::mm(0.0), Length::mm(0.4), Length::mm(0.0), Length::mm(0.4));

    for (ri, row) in rows.iter().enumerate() {
        if let Some(mut tr) = table.row(ri) {
            tr.set_height_exact(Length::mm(cell));
            tr.set_cant_split();
        }
        if has_labels {
            if let Some(mut tc) = table.cell(ri, 0) {
                tc.set_width(Length::mm(label_w));
                tc.set_vertical_alignment(VerticalAlignment::Center);
                let mut p = tc.paragraph_mut(0).expect("fresh cell has a paragraph");
                spacing(&mut p, 0.0, 0.0);
                p.set_alignment(Alignment::Right);
                p.set_indent_right(Length::mm(1.2));
                let label = row_labels.and_then(|l| l.get(ri)).map(String::as_str).unwrap_or("");
                styled_run(&mut p, label, &RunStyle::body().with_size(t::PT_SMALL).color(t::MUTED));
            }
        }
        for (ci, square) in row.iter().enumerate() {
            let Some(mut tc) = table.cell(ri, ci + usize::from(has_labels)) else { continue };
            tc.set_width(Length::mm(cell));
            tc.set_vertical_alignment(VerticalAlignment::Center);
            let Some(square) = square else { continue };
            let blocked = square.blocked.unwrap_or(false);
            let mark = square.mark.unwrap_or(false);
            tc.set_borders(BorderStyle::Single, 6, if mark { ctx.accent.as_str() } else { t::MUTED });
            if blocked {
                tc.set_shading(t::INK);
            } else if mark {
                tc.set_shading(t::SOFT);
            }
            let mut p = tc.paragraph_mut(0).expect("fresh cell has a paragraph");
            spacing(&mut p, 0.0, 0.0);
            p.set_alignment(Alignment::Center);
            let ch = square.char.as_deref().unwrap_or("").trim();
            let style = if mark { RunStyle::body().with_size(letter_pt).bold().color(&ctx.accent) } else { RunStyle::body().with_size(letter_pt).bold() };
            // An empty run still sizes the line so blank squares keep the row pitch.
            styled_run(&mut p, ch, &style);
        }
    }
    gap(sink, 4.0);
}

fn uppercase_inline(inline: &Inline) -> Inline {
    match inline {
        Inline::Text { text, bold, italic, underline, strike, code, sup, sub, tone } => Inline::Text {
            text: text.to_uppercase(),
            bold: *bold,
            italic: *italic,
            underline: *underline,
            strike: *strike,
            code: *code,
            sup: *sup,
            sub: *sub,
            tone: *tone,
        },
        other => other.clone(),
    }
}

fn run_style_for(style: ParagraphStyle) -> RunStyle {
    match style {
        ParagraphStyle::Body => RunStyle::body(),
        ParagraphStyle::Lead => RunStyle::body().with_size(t::PT_LEAD),
        ParagraphStyle::Small => RunStyle::body().with_size(t::PT_SMALL),
        ParagraphStyle::Caption => RunStyle::body().with_size(t::PT_CAPTION).italic().color(t::MUTED),
        ParagraphStyle::Label => RunStyle::body().with_size(t::PT_LABEL).bold().color(t::MUTED),
    }
}

fn apply_paragraph_style(p: &mut Paragraph<'_>, frame: &Frame, style: ParagraphStyle) {
    let size = run_style_for(style).size;
    match style {
        ParagraphStyle::Body => spacing_sized(p, 0.0, 4.0, size, t::LINE_MULTIPLE),
        ParagraphStyle::Lead => spacing_sized(p, 0.0, 6.0, size, t::LINE_MULTIPLE),
        ParagraphStyle::Small => spacing_sized(p, 0.0, 3.0, size, t::LINE_MULTIPLE),
        ParagraphStyle::Caption => {
            spacing_sized(p, 2.0, 6.0, size, t::LINE_MULTIPLE);
            p.set_alignment(Alignment::Center);
        }
        ParagraphStyle::Label => spacing_sized(p, 0.0, 1.0, size, 1.0),
    }
    if frame.indent > 0.0 {
        p.set_indent_left(Length::mm(frame.indent));
    }
}

// ---------------------------------------------------------------------------
// Headings
// ---------------------------------------------------------------------------

#[allow(clippy::too_many_arguments)]
fn render_heading<S: Sink>(
    sink: &mut S,
    ctx: &mut Ctx,
    frame: &Frame,
    level: u8,
    content: &[Inline],
    kicker: Option<&str>,
    meta: Option<&str>,
    page_break_before: bool,
) {
    let mut first = true;
    if let Some(kicker) = kicker.filter(|k| !k.trim().is_empty()) {
        let mut p = sink.para();
        if page_break_before {
            p.set_page_break_before(true);
        }
        first = false;
        spacing(&mut p, 0.0, 1.0);
        p.set_keep_with_next(true);
        if frame.indent > 0.0 {
            p.set_indent_left(Length::mm(frame.indent));
        }
        let style = RunStyle::body().with_size(t::PT_KICKER).bold().color(t::MUTED);
        let mut run = styled_run(&mut p, &kicker.to_uppercase(), &style);
        run.set_character_spacing(Length::pt(0.8));
    }

    let mut p = sink.para();
    if first && page_break_before {
        p.set_page_break_before(true);
    }
    p.set_keep_with_next(true);
    p.set_keep_together(true);
    if frame.indent > 0.0 {
        p.set_indent_left(Length::mm(frame.indent));
    }
    let style = match level {
        1 => {
            p.set_space_before(Length::pt(0.0));
            p.set_space_after(Length::pt(6.0));
            p.set_line_spacing_at_least(t::line_pitch_pt(t::PT_H1, 1.0));
            RunStyle::body().with_size(t::PT_H1).bold()
        }
        2 => {
            p.set_space_before(Length::pt(if kicker.is_some() { 0.0 } else { 14.0 }));
            p.set_space_after(Length::pt(4.0));
            p.set_line_spacing_at_least(t::line_pitch_pt(t::PT_H2, 1.0));
            p.set_border_bottom(BorderStyle::Single, 4, t::RULE);
            RunStyle::body().with_size(t::PT_H2).bold().color(&ctx.accent)
        }
        _ => {
            p.set_space_before(Length::pt(9.0));
            p.set_space_after(Length::pt(2.5));
            p.set_line_spacing_at_least(t::line_pitch_pt(t::PT_H3, 1.0));
            RunStyle::body().with_size(t::PT_H3).bold()
        }
    };
    write_inlines(ctx, &mut p, content, &style);
    if let Some(meta) = meta.filter(|m| !m.trim().is_empty()) {
        p.set_add_tab_stop(TabAlignment::Right, Length::mm(frame.width));
        tab(&mut p, &style);
        let meta_style = RunStyle::body().with_size(if level == 1 { t::PT_BODY } else { t::PT_META }).color(t::MUTED);
        styled_run(&mut p, meta, &meta_style);
    }
}

// ---------------------------------------------------------------------------
// Lists
// ---------------------------------------------------------------------------

fn render_list<S: Sink>(sink: &mut S, ctx: &mut Ctx, frame: &Frame, ordered: bool, items: &[ListItem], tight: bool) {
    let gutter = t::LIST_INDENT_MM;
    for (idx, item) in items.iter().enumerate() {
        let marker = if ordered {
            ordered_marker(frame.list_depth, item.number.unwrap_or(idx as i64 + 1))
        } else {
            match frame.list_depth % 2 {
                0 => "•",
                _ => "–",
            }
            .to_string()
        };
        let mut p = sink.para();
        spacing(&mut p, 0.0, if tight { 1.5 } else { 3.0 });
        p.set_indent_left(Length::mm(frame.indent + gutter));
        p.set_hanging_indent(Length::mm(gutter));
        p.set_add_tab_stop(TabAlignment::Left, Length::mm(frame.indent + gutter));
        let body = RunStyle::body();
        let marker_style = if ordered { body.clone().bold().color(&ctx.accent) } else { body.clone().color(&ctx.accent) };
        styled_run(&mut p, &marker, &marker_style);
        tab(&mut p, &body);
        write_inlines(ctx, &mut p, &item.content, &body);
        if let Some(children) = &item.children {
            render_blocks(sink, ctx, &frame.nested_list(gutter), children);
        }
    }
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

fn column_widths(total: f64, cols: usize, weights: Option<&[f64]>) -> Vec<f64> {
    let weights: Vec<f64> = match weights {
        Some(w) if w.len() == cols && w.iter().all(|x| *x > 0.0) => w.to_vec(),
        _ => vec![1.0; cols],
    };
    let sum: f64 = weights.iter().sum();
    weights.iter().map(|w| total * w / sum).collect()
}

/// Content-driven column weights: the longest single word (which cannot
/// wrap) sets a floor, the average line length adds a share, so label
/// columns get room and numeric columns stay compact.
fn auto_column_weights(rows: &[Vec<TableCell>], cols: usize) -> Option<Vec<f64>> {
    if cols < 2 {
        return None;
    }
    let mut longest_word = vec![0usize; cols];
    let mut total_chars = vec![0usize; cols];
    for row in rows {
        let mut col = 0usize;
        for cell in row {
            let span = cell.col_span.unwrap_or(1).max(1) as usize;
            if span == 1 && col < cols {
                let text = cell_plain_text(&cell.blocks);
                let word = text.split_whitespace().map(|w| w.chars().count()).max().unwrap_or(0);
                longest_word[col] = longest_word[col].max(word);
                total_chars[col] = total_chars[col].max(text.chars().count());
            }
            col += span;
        }
    }
    let weights: Vec<f64> = (0..cols)
        .map(|c| {
            let floor = (longest_word[c] as f64).clamp(3.0, 22.0);
            let body = (total_chars[c] as f64).sqrt().clamp(2.0, 9.0);
            floor + body
        })
        .collect();
    Some(weights)
}

fn cell_plain_text(blocks: &[Block]) -> String {
    let mut out = String::new();
    for block in blocks {
        match block {
            Block::Paragraph { content, .. } | Block::Heading { content, .. } => {
                for inline in content {
                    match inline {
                        Inline::Text { text, .. } | Inline::Link { text, .. } => out.push_str(text),
                        Inline::Math { tex, .. } => out.push_str(tex),
                        Inline::Blank { .. } => out.push_str("________"),
                        Inline::Break => out.push(' '),
                    }
                }
                out.push(' ');
            }
            Block::List { items, .. } => {
                for item in items {
                    out.push_str(&cell_plain_text(&[Block::Paragraph {
                        content: item.content.clone(),
                        style: None,
                        align: None,
                        trailing: None,
                        indent: None,
                        keep_with_next: None,
                    }]));
                }
            }
            _ => out.push_str("          "),
        }
    }
    out
}

#[allow(clippy::too_many_arguments)]
fn render_table<S: Sink>(
    sink: &mut S,
    ctx: &mut Ctx,
    frame: &Frame,
    rows: &[Vec<TableCell>],
    columns: Option<&[f64]>,
    header_rows: usize,
    style: TableStyle,
    dense: bool,
) {
    let cols = rows.iter().map(|r| r.iter().map(|c| c.col_span.unwrap_or(1).max(1) as usize).sum::<usize>()).max().unwrap_or(0);
    if cols == 0 || rows.is_empty() {
        return;
    }
    let total = frame.inner_width();
    let auto = auto_column_weights(rows, cols);
    let widths = column_widths(total, cols, columns.or(auto.as_deref()));
    let (pad_v, pad_h) = if dense { (0.8, 1.8) } else { (1.4, 2.2) };

    let mut table = sink.table(rows.len(), cols);
    table.set_width(Length::mm(total));
    table.set_layout_fixed();
    if frame.indent > 0.0 {
        table.set_indent(Length::mm(frame.indent));
    }
    for (i, w) in widths.iter().enumerate() {
        table.set_column_width(i, Length::mm(*w));
    }
    match style {
        TableStyle::Grid | TableStyle::Rules => table.set_borders(BorderStyle::Single, 4, t::RULE),
        TableStyle::Plain => table.set_borders(BorderStyle::None, 0, "auto"),
    }
    table.set_cell_margins(Length::mm(pad_v), Length::mm(pad_h), Length::mm(pad_v), Length::mm(pad_h));

    for (ri, row_cells) in rows.iter().enumerate() {
        let is_header = ri < header_rows;
        if let Some(mut row) = table.row(ri) {
            if is_header {
                row.set_header();
            }
            row.set_cant_split();
        }
        let mut col = 0usize;
        for (cell_index, cell_ir) in row_cells.iter().enumerate() {
            let span = cell_ir.col_span.unwrap_or(1).max(1) as usize;
            // A grid span consumes physical cells. Later cells are indexed by
            // their position in the row, while widths use logical grid columns.
            if span > 1 {
                table.set_cell_grid_span_checked(ri, cell_index, Some(span as u32))
                    .expect("fresh table has enough empty cells for the computed spans");
            }
            let Some(mut cell) = table.cell(ri, cell_index) else { break };
            let cell_w: f64 = widths[col..col + span].iter().sum();
            cell.set_width(Length::mm(cell_w));
            if is_header || cell_ir.shade.unwrap_or(false) {
                cell.set_shading(t::SOFT);
            }
            cell.set_vertical_alignment(match cell_ir.v_align.unwrap_or_default() {
                VAlign::Top => VerticalAlignment::Top,
                VAlign::Middle => VerticalAlignment::Center,
                VAlign::Bottom => VerticalAlignment::Bottom,
            });
            let inner = Frame::new(cell_w - 2.0 * pad_h);
            render_cell_blocks(&mut cell, ctx, &inner, &cell_ir.blocks, is_header, cell_ir.align);
            col += span;
        }
    }
    keep_table_rows(&mut table, rows.len(), cols, ctx.keep);
    gap(sink, 4.0);
}

/// Cell content with table-specific tweaks: compact paragraph spacing, bold
/// header text, optional horizontal alignment.
fn render_cell_blocks(cell: &mut Cell<'_>, ctx: &mut Ctx, frame: &Frame, blocks: &[Block], header: bool, align: Option<Align>) {
    let before = cell.paragraph_count();
    for block in blocks {
        match block {
            Block::Paragraph { content, style, align: p_align, trailing, indent, keep_with_next: _ } => {
                let style = style.unwrap_or_default();
                let mut p = cell.para();
                spacing(&mut p, 0.0, 1.5);
                let extra = indent.unwrap_or(0) as f64 * t::LIST_INDENT_MM;
                if extra > 0.0 {
                    p.set_indent_left(Length::mm(extra));
                }
                set_align(&mut p, p_align.or(align));
                let mut rs = run_style_for(style);
                if header {
                    rs.bold = true;
                }
                if style == ParagraphStyle::Body {
                    rs.size = t::PT_BODY - 0.5;
                }
                if style == ParagraphStyle::Label {
                    let upper: Vec<Inline> = content.iter().map(uppercase_inline).collect();
                    write_inlines(ctx, &mut p, &upper, &rs);
                } else {
                    write_inlines(ctx, &mut p, content, &rs);
                }
                if let Some(tr) = trailing.as_ref().filter(|tr| !tr.is_empty()) {
                    p.set_add_tab_stop(TabAlignment::Right, Length::mm(frame.width));
                    tab(&mut p, &rs);
                    write_inlines(ctx, &mut p, tr, &rs);
                }
            }
            other => render_block(cell, ctx, frame, other),
        }
    }
    if cell.paragraph_count() > before {
        cell.remove_first_empty_paragraph();
    }
}

// ---------------------------------------------------------------------------
// Panels
// ---------------------------------------------------------------------------

fn render_panel<S: Sink>(sink: &mut S, ctx: &mut Ctx, frame: &Frame, variant: PanelVariant, label: Option<&str>, blocks: &[Block]) {
    let colors = panel_colors(variant);
    let total = frame.inner_width();
    let bar = t::PANEL_BAR_MM;
    let pad_h = 3.5;
    let pad_v = 2.4;
    let body_w = total - bar;

    // One table row per content unit (paragraph, list item, …) so both Word
    // and the PDF layout engine can break a long panel between units instead
    // of pushing the whole card to the next page.
    let units = panel_units(blocks);
    if units.is_empty() {
        return;
    }
    let label = label.filter(|l| !l.trim().is_empty());
    let row_count = units.len();

    gap(sink, 2.0);
    let mut table = sink.table(row_count, 2);
    table.set_width(Length::mm(total));
    table.set_layout_fixed();
    if frame.indent > 0.0 {
        table.set_indent(Length::mm(frame.indent));
    }
    table.set_column_width(0, Length::mm(bar));
    table.set_column_width(1, Length::mm(body_w));
    table.set_borders(BorderStyle::None, 0, "auto");
    table.set_cell_margins(Length::mm(0.0), Length::mm(pad_h), Length::mm(0.0), Length::mm(pad_h));

    let inner = Frame::new(body_w - 2.0 * pad_h);
    let last = row_count - 1;
    for (row_idx, unit) in units.iter().enumerate() {
        panel_bar_cell(&mut table, row_idx, bar, colors.bar);
        if let Some(mut row) = table.row(row_idx) {
            if estimate_chars(unit) <= 320 {
                row.set_cant_split();
            }
        }
        if let Some(mut cell) = table.cell(row_idx, 1) {
            cell.set_width(Length::mm(body_w));
            cell.set_shading(colors.fill);
            // The label shares the first row with the first unit so it can
            // never be orphaned at the bottom of a page.
            if row_idx == 0 {
                if let Some(label) = label {
                    let mut p = cell.paragraph_mut(0).expect("fresh cell has a paragraph");
                    spacing(&mut p, pad_v, 1.5);
                    p.set_keep_with_next(true);
                    let style = RunStyle::body().with_size(t::PT_LABEL).bold().color(colors.bar);
                    let mut run = styled_run(&mut p, &label.to_uppercase(), &style);
                    run.set_character_spacing(Length::pt(0.7));
                }
            }
            render_blocks(&mut cell, ctx, &inner, unit);
            cell.remove_first_empty_paragraph();
            if row_idx == 0 && label.is_none() {
                if let Some(mut p) = cell.paragraph_mut(0) {
                    p.set_space_before(Length::pt(pad_v * 72.0 / 25.4));
                }
            }
            if row_idx == last {
                let n = cell.paragraph_count();
                if let Some(mut p) = cell.paragraph_mut(n.saturating_sub(1)) {
                    p.set_space_after(Length::pt(pad_v * 72.0 / 25.4));
                }
            }
        }
    }
    gap(sink, 5.0);
}

/// Shaded accent bar cell of a panel row (empty 1pt paragraph, zero spacing).
fn panel_bar_cell(table: &mut Table<'_>, row: usize, bar: f64, color: &str) {
    if let Some(mut bar_cell) = table.cell(row, 0) {
        bar_cell.set_width(Length::mm(bar));
        bar_cell.set_shading(color);
        let mut p = bar_cell.paragraph_mut(0).expect("fresh cell has a paragraph");
        p.set_space_before(Length::pt(0.0));
        p.set_space_after(Length::pt(0.0));
        p.set_line_spacing(1.0);
        let mut r = p.add_run("");
        r.set_size(1.0);
    }
}

/// Split panel content into breakable units: every block is a unit, except
/// lists which contribute one unit per top-level item (numbering preserved).
fn panel_units(blocks: &[Block]) -> Vec<Vec<Block>> {
    let mut units: Vec<Vec<Block>> = Vec::new();
    for block in blocks {
        match block {
            Block::List { ordered, items, tight } => {
                for (idx, item) in items.iter().enumerate() {
                    let mut single = item.clone();
                    single.number = Some(item.number.unwrap_or(idx as i64 + 1));
                    units.push(vec![Block::List { ordered: *ordered, items: vec![single], tight: *tight }]);
                }
            }
            other => units.push(vec![other.clone()]),
        }
    }
    units
}

/// Rough content size (characters, images count as a paragraph's worth) used
/// for keep-together heuristics.
fn estimate_chars(blocks: &[Block]) -> usize {
    fn inl(inlines: &[Inline]) -> usize {
        inlines
            .iter()
            .map(|i| match i {
                Inline::Text { text, .. } => text.chars().count(),
                Inline::Math { tex, .. } => tex.chars().count(),
                Inline::Link { text, .. } => text.chars().count(),
                Inline::Blank { .. } => 12,
                Inline::Break => 40,
            })
            .sum()
    }
    fn items(list: &[ListItem]) -> usize {
        list.iter().map(|it| 20 + inl(&it.content) + it.children.as_deref().map(estimate_chars).unwrap_or(0)).sum()
    }
    blocks
        .iter()
        .map(|b| match b {
            Block::Paragraph { content, .. } => 20 + inl(content),
            Block::Heading { content, .. } => 40 + inl(content),
            Block::List { items: list, .. } => items(list),
            Block::Table { rows, .. } => rows.iter().map(|r| 40 + r.iter().map(|c| estimate_chars(&c.blocks)).sum::<usize>()).sum(),
            Block::Panel { blocks, .. } => 40 + estimate_chars(blocks),
            Block::Image { .. } | Block::Gallery { .. } => 600,
            Block::Facts { rows } => rows.len() * 60,
            Block::AnswerLines { lines } => *lines as usize * 90,
            Block::AnswerBox { .. } => 400,
            Block::Fields { .. } => 80,
            Block::Task { blocks, .. } => 30 + estimate_chars(blocks),
            Block::Options { items: opts, .. } => opts.iter().map(|o| 10 + inl(&o.content)).sum(),
            Block::Cards { items: cards, .. } => cards.iter().map(|c| 60 + c.parts.iter().map(|p| estimate_chars(&p.blocks)).sum::<usize>()).sum(),
            Block::Grid { rows, .. } => rows.iter().map(|r| 10 + r.len() * 8).sum(),
            Block::Rule { .. } | Block::Spacer { .. } | Block::PageBreak => 10,
        })
        .sum()
}

/// Last paragraph in a shaded cell should not carry its bottom spacing.
#[allow(dead_code)]
fn trim_trailing_space(cell: &mut Cell<'_>) {
    let n = cell.paragraph_count();
    if n == 0 {
        return;
    }
    if let Some(mut p) = cell.paragraph_mut(n - 1) {
        p.set_space_after(Length::pt(0.0));
    }
}

// ---------------------------------------------------------------------------
// Facts (key / value card)
// ---------------------------------------------------------------------------

fn render_facts<S: Sink>(sink: &mut S, ctx: &mut Ctx, frame: &Frame, rows: &[crate::ir::FactRow]) {
    if rows.is_empty() {
        return;
    }
    let total = frame.inner_width();
    let label_w = 34.0;
    let value_w = total - label_w;
    let pad_h = 2.4;
    let pad_v = 1.8;

    let mut table = sink.table(rows.len(), 2);
    table.set_width(Length::mm(total));
    table.set_layout_fixed();
    if frame.indent > 0.0 {
        table.set_indent(Length::mm(frame.indent));
    }
    table.set_column_width(0, Length::mm(label_w));
    table.set_column_width(1, Length::mm(value_w));
    table.set_borders(BorderStyle::Single, 4, t::RULE);
    table.set_cell_margins(Length::mm(pad_v), Length::mm(pad_h), Length::mm(pad_v), Length::mm(pad_h));

    for (ri, row) in rows.iter().enumerate() {
        if let Some(mut r) = table.row(ri) {
            r.set_cant_split();
        }
        if let Some(mut cell) = table.cell(ri, 0) {
            cell.set_width(Length::mm(label_w));
            cell.set_shading(t::SOFT);
            let mut p = cell.paragraph_mut(0).expect("fresh cell has a paragraph");
            spacing(&mut p, 1.0, 0.0);
            let style = RunStyle::body().with_size(t::PT_LABEL).bold().color(t::MUTED);
            let mut run = styled_run(&mut p, &row.label.to_uppercase(), &style);
            run.set_character_spacing(Length::pt(0.5));
        }
        if let Some(mut cell) = table.cell(ri, 1) {
            cell.set_width(Length::mm(value_w));
            let inner = Frame::new(value_w - 2.0 * pad_h);
            render_cell_blocks(&mut cell, ctx, &inner, &row.blocks, false, None);
            trim_trailing_space(&mut cell);
        }
    }
    gap(sink, 6.0);
}

// ---------------------------------------------------------------------------
// Images
// ---------------------------------------------------------------------------

/// Fit an image into the frame: never wider than `max_width` of the frame,
/// never taller than `max_height_mm`, never upscaled past ~110 dpi.
fn fit_image(img: &EmbeddedImage, frame_w: f64, max_width: f64, max_height_mm: f64) -> (f64, f64) {
    let aspect = img.height_px.max(1) as f64 / img.width_px.max(1) as f64;
    let natural_w = img.width_px as f64 * 25.4 / 110.0;
    let mut w = (frame_w * max_width).min(natural_w.max(45.0)).min(frame_w);
    let mut h = w * aspect;
    if h > max_height_mm {
        h = max_height_mm;
        w = h / aspect;
    }
    (w, h)
}

fn render_image<S: Sink>(sink: &mut S, ctx: &mut Ctx, frame: &Frame, image: &Image, align: Option<Align>, max_h_override: Option<f64>) {
    let key = image_key(image);
    let Some(img) = ctx.images.get(&key) else {
        return;
    };
    let (w, h) = fit_image(img, frame.inner_width(), image.max_width.unwrap_or(1.0), max_h_override.unwrap_or(image.max_height_mm.unwrap_or(72.0)));
    let has_caption = image.caption.as_deref().map(|c| !c.trim().is_empty()).unwrap_or(false);
    {
        let mut p = sink.picture(img, Length::mm(w), Length::mm(h));
        p.set_alignment(match align {
            Some(Align::Left) => Alignment::Left,
            Some(Align::Right) => Alignment::Right,
            _ => Alignment::Center,
        });
        p.set_space_before(Length::pt(4.0));
        p.set_space_after(Length::pt(if has_caption { 1.0 } else { 6.0 }));
        p.set_keep_together(true);
        if has_caption {
            p.set_keep_with_next(true);
        }
        if frame.indent > 0.0 {
            p.set_indent_left(Length::mm(frame.indent));
        }
    }
    if let Some(caption) = image.caption.as_deref().filter(|c| !c.trim().is_empty()) {
        let mut p = sink.para();
        apply_paragraph_style(&mut p, frame, ParagraphStyle::Caption);
        write_inlines(
            ctx,
            &mut p,
            &[Inline::Text { text: caption.trim().to_string(), bold: None, italic: None, underline: None, strike: None, code: None, sup: None, sub: None, tone: None }],
            &run_style_for(ParagraphStyle::Caption),
        );
    }
}

fn render_gallery<S: Sink>(sink: &mut S, ctx: &mut Ctx, frame: &Frame, images: &[Image]) {
    let present: Vec<&Image> = images.iter().filter(|i| ctx.images.contains_key(&image_key(i))).collect();
    if present.is_empty() {
        return;
    }
    if present.len() == 1 {
        render_image(sink, ctx, frame, present[0], Some(Align::Center), Some(85.0));
        return;
    }
    let per_row = present.len().min(3);
    let total = frame.inner_width();
    let pad_h = 1.5;
    let col_w = total / per_row as f64;
    for chunk in present.chunks(per_row) {
        let mut table = sink.table(1, per_row);
        table.set_width(Length::mm(total));
        table.set_layout_fixed();
        if frame.indent > 0.0 {
            table.set_indent(Length::mm(frame.indent));
        }
        for i in 0..per_row {
            table.set_column_width(i, Length::mm(col_w));
        }
        table.set_borders(BorderStyle::None, 0, "auto");
        table.set_cell_margins(Length::mm(0.5), Length::mm(pad_h), Length::mm(0.5), Length::mm(pad_h));
        if let Some(mut row) = table.row(0) {
            row.set_cant_split();
        }
        for (i, image) in chunk.iter().enumerate() {
            let Some(mut cell) = table.cell(0, i) else { continue };
            cell.set_width(Length::mm(col_w));
            let inner = Frame::new(col_w - 2.0 * pad_h);
            let max_h = if per_row == 2 { 70.0 } else { 55.0 };
            render_image(&mut cell, ctx, &inner, image, Some(Align::Center), Some(max_h));
            cell.remove_first_empty_paragraph();
        }
    }
    gap(sink, 3.0);
}

// ---------------------------------------------------------------------------
// Worksheet primitives
// ---------------------------------------------------------------------------

fn render_fields<S: Sink>(sink: &mut S, frame: &Frame, items: &[Field]) {
    if items.is_empty() {
        return;
    }
    let total = frame.inner_width();
    let label_style = RunStyle::body().with_size(t::PT_SMALL).color(t::MUTED);
    // Estimated natural widths → scale to fill the line.
    let natural: Vec<f64> = items
        .iter()
        .map(|f| f.label.chars().count() as f64 * t::char_width_mm(t::PT_SMALL) + 3.0 + f.width_mm + f.value.as_deref().map(|v| v.chars().count() as f64 * t::char_width_mm(t::PT_BODY)).unwrap_or(0.0) + 4.0)
        .collect();
    let sum: f64 = natural.iter().sum();
    let scale = if sum > total { total / sum } else { 1.0 };
    let widths: Vec<f64> = natural.iter().map(|w| w * scale).collect();
    let used: f64 = widths.iter().sum();

    let cols = items.len();
    let mut table = sink.table(1, cols);
    table.set_width(Length::mm(used.min(total)));
    table.set_layout_fixed();
    if frame.indent > 0.0 {
        table.set_indent(Length::mm(frame.indent));
    }
    for (i, w) in widths.iter().enumerate() {
        table.set_column_width(i, Length::mm(*w));
    }
    table.set_borders(BorderStyle::None, 0, "auto");
    table.set_cell_margins(Length::mm(1.0), Length::mm(0.0), Length::mm(1.0), Length::mm(0.0));
    for (i, field) in items.iter().enumerate() {
        let Some(mut cell) = table.cell(0, i) else { continue };
        cell.set_width(Length::mm(widths[i]));
        cell.set_vertical_alignment(VerticalAlignment::Bottom);
        let mut p = cell.paragraph_mut(0).expect("fresh cell has a paragraph");
        spacing(&mut p, 4.0, 0.0);
        styled_run(&mut p, &format!("{}: ", field.label), &label_style);
        let blank_w = field.width_mm * scale;
        let underscores = (blank_w / (t::char_width_mm(t::PT_BODY) * 1.0)).round().max(4.0) as usize;
        styled_run(&mut p, &"_".repeat(underscores), &RunStyle::body().color(t::BLANK_LINE));
        if let Some(v) = field.value.as_deref() {
            styled_run(&mut p, v, &RunStyle::body().bold());
        }
    }
    gap(sink, 4.0);
}

fn render_task<S: Sink>(sink: &mut S, ctx: &mut Ctx, frame: &Frame, number: Option<i64>, points: Option<&str>, prompt: &[Inline], blocks: &[Block]) {
    let gutter = t::TASK_GUTTER_MM;
    let inner = frame.indented(gutter);
    let mut body_blocks: &[Block] = blocks;

    // Prompt line (or, when there is no prompt, the first paragraph of the body).
    let mut p = sink.para();
    p.set_space_before(Length::pt(9.0));
    p.set_space_after(Length::pt(3.0));
    p.set_line_spacing_at_least(t::line_pitch_pt(t::PT_BODY, t::LINE_MULTIPLE));
    p.set_indent_left(Length::mm(frame.indent + gutter));
    p.set_hanging_indent(Length::mm(gutter));
    p.set_add_tab_stop(TabAlignment::Left, Length::mm(frame.indent + gutter));
    p.set_keep_with_next(true);
    p.set_keep_together(true);
    let num_style = RunStyle::body().with_size(t::PT_BODY + 0.5).bold().color(&ctx.accent);
    if let Some(n) = number {
        styled_run(&mut p, &format!("{n}."), &num_style);
    }
    tab(&mut p, &RunStyle::body());
    let prompt_style = RunStyle::body().bold();
    let mut trailing: Option<&[Inline]> = None;
    if !prompt.is_empty() {
        write_inlines(ctx, &mut p, prompt, &prompt_style);
    } else if let Some((Block::Paragraph { content, trailing: tr, .. }, rest)) = blocks.split_first() {
        write_inlines(ctx, &mut p, content, &RunStyle::body());
        trailing = tr.as_deref().filter(|x| !x.is_empty());
        body_blocks = rest;
    }
    if points.is_some() || trailing.is_some() {
        p.set_add_tab_stop(TabAlignment::Right, Length::mm(frame.width));
        tab(&mut p, &RunStyle::body());
        if let Some(tr) = trailing {
            write_inlines(ctx, &mut p, tr, &RunStyle::body());
            if points.is_some() {
                styled_run(&mut p, "    ", &RunStyle::body());
            }
        }
        if let Some(pts) = points {
            styled_run(&mut p, pts, &RunStyle::body().with_size(t::PT_META).color(t::MUTED));
        }
    }
    drop(p);

    // Short tasks never straddle a page: chain their paragraphs (and table
    // rows) with keep-with-next, leaving the very last one free so the chain
    // does not swallow the following task.
    let chain = estimate_chars(body_blocks) <= t::KEEP_TASK_CHARS;
    let last_index = body_blocks.len().saturating_sub(1);
    for (index, block) in body_blocks.iter().enumerate() {
        ctx.keep = match (chain, index == last_index) {
            (false, _) => Keep::Off,
            (true, false) => Keep::All,
            (true, true) => Keep::ExceptLast,
        };
        let start = sink.direct_paragraph_count();
        render_block(sink, ctx, &inner, block);
        let end = sink.direct_paragraph_count();
        let limit = match ctx.keep {
            Keep::Off => start,
            Keep::All => end,
            Keep::ExceptLast => end.saturating_sub(1),
        };
        for i in start..limit {
            if let Some(mut para) = sink.direct_paragraph(i) {
                para.set_keep_with_next(true);
            }
        }
    }
    ctx.keep = Keep::Off;
}

fn render_options<S: Sink>(sink: &mut S, ctx: &mut Ctx, frame: &Frame, columns: Option<u8>, items: &[OptionItem], reveal: bool) {
    let gutter = t::OPTION_GUTTER_MM;
    let cols = columns.unwrap_or(1).clamp(1, 2) as usize;
    let has_images = items.iter().any(|i| i.image.is_some());
    if cols == 1 || has_images {
        for item in items {
            write_option(sink, ctx, frame, gutter, item, reveal);
        }
        gap(sink, 2.0);
        return;
    }
    let total = frame.inner_width();
    let col_w = total / 2.0;
    let rows = items.len().div_ceil(2);
    let mut table = sink.table(rows, 2);
    table.set_width(Length::mm(total));
    table.set_layout_fixed();
    if frame.indent > 0.0 {
        table.set_indent(Length::mm(frame.indent));
    }
    table.set_column_width(0, Length::mm(col_w));
    table.set_column_width(1, Length::mm(col_w));
    table.set_borders(BorderStyle::None, 0, "auto");
    table.set_cell_margins(Length::mm(0.0), Length::mm(0.0), Length::mm(0.0), Length::mm(0.0));
    for (i, item) in items.iter().enumerate() {
        let (r, c) = (i / 2, i % 2);
        let Some(mut cell) = table.cell(r, c) else { continue };
        cell.set_width(Length::mm(col_w));
        let inner = Frame::new(col_w);
        write_option(&mut cell, ctx, &inner, gutter, item, reveal);
        cell.remove_first_empty_paragraph();
    }
    keep_table_rows(&mut table, rows, 2, ctx.keep);
    gap(sink, 3.0);
}

fn write_option<S: Sink>(sink: &mut S, ctx: &mut Ctx, frame: &Frame, gutter: f64, item: &OptionItem, reveal: bool) {
    let correct = reveal && item.correct.unwrap_or(false);
    let mut p = sink.para();
    spacing(&mut p, 0.0, 2.0);
    p.set_indent_left(Length::mm(frame.indent + gutter));
    p.set_hanging_indent(Length::mm(gutter));
    p.set_add_tab_stop(TabAlignment::Left, Length::mm(frame.indent + gutter));
    let label_style = if correct {
        RunStyle::body().bold().color(t::ANSWER)
    } else {
        RunStyle::body().bold().color(&ctx.accent)
    };
    styled_run(&mut p, &format!("{})", item.label), &label_style);
    tab(&mut p, &RunStyle::body());
    let text_style = if correct { RunStyle::body().bold().color(t::ANSWER) } else { RunStyle::body() };
    write_inlines(ctx, &mut p, &item.content, &text_style);
    if let Some(image) = &item.image {
        let f = frame.indented(gutter);
        render_image(sink, ctx, &f, image, Some(Align::Left), Some(45.0));
    }
}

// ---------------------------------------------------------------------------
// Header / footer (raw WordprocessingML so we get PAGE / NUMPAGES fields)
// ---------------------------------------------------------------------------

fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;").replace('"', "&quot;")
}

fn rpr(size_half_pt: u32, color: &str, bold: bool) -> String {
    format!(
        "<w:rPr><w:rFonts w:ascii=\"{f}\" w:hAnsi=\"{f}\" w:cs=\"{f}\"/>{b}<w:color w:val=\"{color}\"/><w:sz w:val=\"{s}\"/><w:szCs w:val=\"{s}\"/></w:rPr>",
        f = t::FONT_BODY,
        b = if bold { "<w:b/>" } else { "" },
        s = size_half_pt
    )
}

fn text_run(text: &str, rpr_xml: &str) -> String {
    format!("<w:r>{rpr_xml}<w:t xml:space=\"preserve\">{}</w:t></w:r>", xml_escape(text))
}

fn field_run(instr: &str, placeholder: &str, rpr_xml: &str) -> String {
    format!(
        "<w:r>{r}<w:fldChar w:fldCharType=\"begin\"/></w:r><w:r>{r}<w:instrText xml:space=\"preserve\"> {instr} </w:instrText></w:r><w:r>{r}<w:fldChar w:fldCharType=\"separate\"/></w:r><w:r>{r}<w:t>{placeholder}</w:t></w:r><w:r>{r}<w:fldChar w:fldCharType=\"end\"/></w:r>",
        r = rpr_xml
    )
}

fn page_number_runs(format: &str, rpr_xml: &str) -> String {
    let mut out = String::new();
    let mut rest = format;
    loop {
        let next_page = rest.find("{page}");
        let next_pages = rest.find("{pages}");
        let (idx, token, instr, ph) = match (next_page, next_pages) {
            (Some(a), Some(b)) if a <= b => (a, "{page}", "PAGE", "1"),
            (Some(a), None) => (a, "{page}", "PAGE", "1"),
            (_, Some(b)) => (b, "{pages}", "NUMPAGES", "1"),
            (None, None) => {
                if !rest.is_empty() {
                    out.push_str(&text_run(rest, rpr_xml));
                }
                break;
            }
        };
        if idx > 0 {
            out.push_str(&text_run(&rest[..idx], rpr_xml));
        }
        out.push_str(&field_run(instr, ph, rpr_xml));
        rest = &rest[idx + token.len()..];
    }
    out
}

fn install_header_footer(doc: &mut Document, ctx: &mut Ctx, ir: &PrintDocument) {
    const NS: &str = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
    let right_tab = (ctx.content_w * 56.6929).round() as i64; // mm → twips
    let muted = rpr(17, t::MUTED, false);

    if let Some(header) = &ir.header {
        let left = header.left.as_deref().unwrap_or("");
        let right = header.right.as_deref().unwrap_or("");
        if !left.is_empty() || !right.is_empty() {
            let xml = format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:hdr {NS}><w:p><w:pPr><w:pBdr><w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"4\" w:color=\"{rule}\"/></w:pBdr><w:tabs><w:tab w:val=\"right\" w:pos=\"{right_tab}\"/></w:tabs><w:spacing w:before=\"0\" w:after=\"0\"/></w:pPr>{l}<w:r>{muted}<w:tab/></w:r>{r}</w:p></w:hdr>",
                rule = t::RULE,
                l = text_run(left, &muted),
                r = text_run(right, &muted),
            );
            doc.set_raw_header_with_images(xml.into_bytes(), &[], HdrFtrType::Default);
        }
    }

    let footer = ir.footer.as_ref();
    let left = footer.and_then(|f| f.left.as_deref()).unwrap_or("");
    let page_numbers = footer.and_then(|f| f.page_numbers).unwrap_or(true);
    let format = footer.and_then(|f| f.page_number_format.as_deref()).unwrap_or("{page} / {pages}");
    let mark = ir.watermark.as_ref().and_then(|w| footer_mark(ctx, w));
    if left.is_empty() && !page_numbers && mark.is_none() {
        return;
    }
    let left_xml = if left.is_empty() { String::new() } else { text_run(left, &muted) };
    let numbers_xml = if page_numbers { page_number_runs(format, &muted) } else { String::new() };

    // `left ⇥ page numbers` on one right tab; with a mark, the numbers move to
    // a second right tab left of the picture so the two never overlap.
    let (tabs_xml, right_xml, images): (String, String, Vec<(&str, &[u8], &str)>) = match &mark {
        None => (
            format!("<w:tab w:val=\"right\" w:pos=\"{right_tab}\"/>"),
            format!("<w:r>{muted}<w:tab/></w:r>{numbers_xml}"),
            Vec::new(),
        ),
        Some(mark) => {
            let numbers_tab = ((ctx.content_w - mark.width_mm - FOOTER_MARK_GAP_MM) * 56.6929).round() as i64;
            let numbers = if page_numbers { format!("<w:r>{muted}<w:tab/></w:r>{numbers_xml}") } else { String::new() };
            (
                format!("<w:tab w:val=\"right\" w:pos=\"{numbers_tab}\"/><w:tab w:val=\"right\" w:pos=\"{right_tab}\"/>"),
                format!("{numbers}<w:r>{muted}<w:tab/></w:r>{}", mark.drawing_xml(&muted)),
                vec![(FOOTER_MARK_REL_ID, mark.bytes.as_slice(), mark.filename)],
            )
        }
    };
    let xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:ftr {NS}><w:p><w:pPr><w:pBdr><w:top w:val=\"single\" w:sz=\"4\" w:space=\"4\" w:color=\"{rule}\"/></w:pBdr><w:tabs>{tabs_xml}</w:tabs><w:spacing w:before=\"0\" w:after=\"0\"/></w:pPr>{left_xml}{right_xml}</w:p></w:ftr>",
        rule = t::RULE,
    );
    doc.set_raw_footer_with_images(xml.into_bytes(), &images, HdrFtrType::Default);
}

/// Relationship id the footer picture is authored with; rdocx remaps on clash.
const FOOTER_MARK_REL_ID: &str = "rIdMark";
/// Space between the page number and the footer mark.
const FOOTER_MARK_GAP_MM: f64 = 5.0;
const FOOTER_MARK_DEFAULT_WIDTH_MM: f64 = 22.0;
const EMU_PER_MM: f64 = 36000.0;

struct FooterMark {
    bytes: Vec<u8>,
    filename: &'static str,
    width_mm: f64,
    height_mm: f64,
    alt: String,
}

impl FooterMark {
    /// Inline picture run; sits on the text baseline of the footer line.
    fn drawing_xml(&self, rpr_xml: &str) -> String {
        let cx = (self.width_mm * EMU_PER_MM).round() as i64;
        let cy = (self.height_mm * EMU_PER_MM).round() as i64;
        format!(
            "<w:r>{rpr_xml}<w:drawing><wp:inline xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">\
<wp:extent cx=\"{cx}\" cy=\"{cy}\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>\
<wp:docPr id=\"1\" name=\"watermark\" descr=\"{alt}\"/>\
<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>\
<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\
<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:nvPicPr><pic:cNvPr id=\"0\" name=\"{file}\"/><pic:cNvPicPr/></pic:nvPicPr>\
<pic:blipFill><a:blip r:embed=\"{rel}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>\
<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>\
</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>",
            alt = xml_escape(&self.alt),
            file = self.filename,
            rel = FOOTER_MARK_REL_ID,
        )
    }
}

/// Decodes and sizes the footer mark; `None` (with a warning) when the image
/// data is unusable, so a broken logo never blocks the document.
fn footer_mark(ctx: &mut Ctx, mark: &crate::ir::Watermark) -> Option<FooterMark> {
    let bytes = match image_bytes(&mark.image) {
        Ok(b) => b,
        Err(e) => {
            ctx.warnings.push(format!("watermark: {e}"));
            return None;
        }
    };
    let Some(info) = oxml_media::probe(&bytes) else {
        ctx.warnings.push(format!("watermark: unsupported or corrupt image ({} bytes, {})", bytes.len(), mark.image.mime));
        return None;
    };
    let filename = match mark.image.mime.as_str() {
        "image/jpeg" => "watermark.jpg",
        "image/gif" => "watermark.gif",
        "image/webp" => "watermark.webp",
        _ => "watermark.png",
    };
    // Never wider than a third of the line: the footer text must keep room.
    let width_mm = mark.width_mm.unwrap_or(FOOTER_MARK_DEFAULT_WIDTH_MM).clamp(5.0, ctx.content_w / 3.0);
    let aspect = info.height_px.max(1) as f64 / info.width_px.max(1) as f64;
    Some(FooterMark {
        bytes,
        filename,
        width_mm,
        height_mm: width_mm * aspect,
        alt: mark.image.alt.clone().unwrap_or_default(),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(json: &str) -> PrintDocument {
        serde_json::from_str(json).expect("valid IR")
    }

    /// A short test sheet: title, meta line, name fields, one closed and one
    /// open task. Small enough for one page.
    fn sample_test() -> PrintDocument {
        parse(
            r##"{
              "version": 1, "locale": "cs", "kind": "test", "title": "Zlomky",
              "footer": { "left": "Zlomky  ·  Skupina A", "pageNumbers": true, "pageNumberFormat": "Strana {page} / {pages}" },
              "blocks": [
                { "kind": "heading", "level": 1, "content": [{ "kind": "text", "text": "Zlomky" }], "kicker": "Test  ·  Matematika", "meta": "Skupina A" },
                { "kind": "paragraph", "style": "small", "content": [{ "kind": "text", "text": "Čas: 15 min", "tone": "muted" }] },
                { "kind": "fields", "items": [{ "label": "Jméno", "widthMm": 52 }, { "label": "Třída", "widthMm": 22 }, { "label": "Body", "widthMm": 22, "value": " / 5" }] },
                { "kind": "rule", "tone": "accent" },
                { "kind": "task", "number": 1, "points": "2 b.", "prompt": [{ "kind": "text", "text": "Kolik je " }, { "kind": "math", "tex": "\\frac{1}{2} + \\frac{1}{4}" }, { "kind": "text", "text": "?" }],
                  "blocks": [{ "kind": "options", "columns": 2, "reveal": false, "items": [
                    { "label": "A", "content": [{ "kind": "math", "tex": "\\frac{2}{6}" }], "correct": false },
                    { "label": "B", "content": [{ "kind": "math", "tex": "\\frac{3}{4}" }], "correct": true },
                    { "label": "C", "content": [{ "kind": "text", "text": "1" }], "correct": false } ] }] },
                { "kind": "task", "number": 2, "points": "3 b.", "prompt": [{ "kind": "text", "text": "Vysvětli, co je smíšené číslo." }],
                  "blocks": [{ "kind": "answerLines", "lines": 3 }] }
              ]
            }"##,
        )
    }


    #[test]
    fn theme_colours_expand_shorthand_and_reject_invalid_ooxml_values() {
        for (input, expected) in [("#abc", "AABBCC"), ("abc", "AABBCC"), (" #12aB34 ", "12AB34")] {
            let mut ir = sample_test();
            ir.theme = Some(crate::ir::Theme { accent: Some(input.to_string()) });
            assert_eq!(Ctx::new(&ir).unwrap().accent, expected);
            let (mut doc, _) = render(&ir).unwrap();
            Document::from_bytes(&doc.to_bytes().unwrap()).unwrap();
        }
        for input in ["", "##123456", "red", "12345", "12345678", "GGGGGG", "🟦"] {
            let mut ir = sample_test();
            ir.theme = Some(crate::ir::Theme { accent: Some(input.to_string()) });
            let Err(error) = render(&ir) else { panic!("accepted invalid colour {input}"); };
            assert!(error.contains("theme accent"), "{error}");
        }
    }

    #[test]
    fn card_assets_are_collected_through_nested_blocks() {
        let ir = parse(&serde_json::json!({
            "version": 1, "locale": "en", "kind": "activity", "title": "Cards",
            "blocks": [{ "kind": "cards", "columns": 1, "items": [{ "parts": [{ "blocks": [
                { "kind": "panel", "variant": "plain", "blocks": [
                    { "kind": "image", "image": { "data": PIXEL_PNG_B64, "mime": "image/png" } }
                ] },
                { "kind": "paragraph", "content": [{ "kind": "link", "href": "https://example.com", "text": "Link" }] }
            ] }] }] }]
        }).to_string());
        let (mut doc, report) = render(&ir).unwrap();
        assert!(report.warnings.is_empty());
        assert_eq!(doc.images().len(), 1);
        let reopened = Document::from_bytes(&doc.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.images().len(), 1);
        let tables = reopened.tables();
        let cell = tables[0].cell(0, 0).unwrap();
        let links: Vec<_> = cell.paragraphs().flat_map(|p| {
            p.hyperlink_spans().into_iter()
                .filter_map(|(_, _, id)| id.and_then(|id| reopened.hyperlink_url(id)))
                .collect::<Vec<_>>()
        }).collect();
        assert_eq!(links, ["https://example.com"]);
    }

    #[test]
    fn table_spans_consume_covered_cells_without_losing_following_content() {
        for spans in [[2, 1], [1, 2], [2, 2], [0, 1]] {
            let row: Vec<_> = spans.iter().enumerate().map(|(i, span)| serde_json::json!({
                "colSpan": span, "blocks": [{ "kind": "paragraph", "content": [{ "kind": "text", "text": format!("cell {i}") }] }]
            })).collect();
            let ir = parse(&serde_json::json!({
                "version": 1, "locale": "en", "kind": "test", "title": "Spans",
                "blocks": [{ "kind": "table", "rows": [row] }]
            }).to_string());
            let (mut doc, _) = render(&ir).unwrap();
            let reopened = Document::from_bytes(&doc.to_bytes().unwrap()).unwrap();
            let tables = reopened.tables();
            let table = &tables[0];
            let row = table.row(0).unwrap();
            assert_eq!(row.cell_count(), 2, "spans {spans:?}");
            assert_eq!(table.column_count(), spans.iter().map(|s| (*s).max(1) as usize).sum::<usize>());
            for (i, span) in spans.iter().enumerate() {
                let cell = row.cell(i).unwrap();
                assert_eq!(cell.grid_span().unwrap_or(1), (*span).max(1));
                assert!(cell.text().contains(&format!("cell {i}")));
            }
        }
    }

    #[test]
    fn renders_a_test_sheet_onto_one_page() {
        let (doc, report) = render(&sample_test()).expect("render");
        assert!(report.warnings.is_empty(), "warnings: {:?}", report.warnings);
        let layout = doc.layout().expect("layout");
        assert_eq!(layout.layout.pages.len(), 1);
        let texts: Vec<String> = doc.paragraphs().iter().map(|p| p.text()).collect();
        assert!(texts.iter().any(|t| t.contains("Zlomky") && t.contains("Skupina A")), "title missing: {texts:?}");
        assert!(texts.iter().any(|t| t.starts_with("1.") && t.contains("2 b.")), "task 1 missing: {texts:?}");
        // Name fields and the two-column options both render as tables.
        let tables = doc.tables();
        assert_eq!(tables.len(), 2, "fields + options");
        let fields: String = (0..tables[0].column_count()).filter_map(|c| tables[0].cell(0, c)).map(|c| c.text()).collect();
        assert!(fields.contains("Jméno") && fields.contains("Body") && fields.contains("/ 5"), "fields: {fields:?}");
        assert_eq!(tables[1].row_count(), 2, "three options in two columns");
        // Office Math is not part of `text()`; the paragraph must carry the equations.
        let prompt = doc.paragraphs().into_iter().find(|p| p.text().starts_with("1.")).unwrap();
        assert_eq!(prompt.equations().count(), 1);
    }

    #[test]
    fn mathml_and_latex_fractions_remain_editable_after_docx_round_trip() {
        let ir = parse(
            r#"{
            "version": 1, "locale": "cs", "kind": "lesson", "title": "Math",
            "blocks": [{ "kind": "paragraph", "content": [
                { "kind": "math", "tex": "\\frac{3}{10}" },
                { "kind": "text", "text": " + " },
                { "kind": "math", "tex": "\\frac{4}{10}", "mathml": "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mn>4</mn><mn>10</mn></mfrac></math>" },
                { "kind": "text", "text": " = " },
                { "kind": "math", "tex": "\\frac{7}{10}", "mathml": "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mn>7</mn><mn>10</mn></mfrac></math>" }
            ] }]
        }"#,
        );
        let (mut doc, report) = render(&ir).expect("render");
        assert!(report.warnings.is_empty(), "{:?}", report.warnings);
        let reopened =
            Document::from_bytes(&doc.to_bytes().expect("DOCX bytes")).expect("reopen DOCX");
        let equation_xml = |document: &Document| -> Vec<String> {
            document
                .paragraphs()
                .iter()
                .flat_map(|p| {
                    p.equations()
                        .map(|eq| {
                            let xml = String::from_utf8(eq.to_xml().expect("OMML")).unwrap();
                            // Reopening retains additional inherited namespace declarations
                            // on oMath. Compare the complete equation body instead.
                            xml.split_once('>').expect("oMath root").1.to_owned()
                        })
                        .collect::<Vec<_>>()
                })
                .collect()
        };
        let original = equation_xml(&doc);
        assert_eq!(original.len(), 3);
        assert_eq!(equation_xml(&reopened), original);
        for xml in original {
            assert!(xml.contains("<m:f>"), "native fraction missing: {xml}");
            assert!(xml.contains("<m:num>") && xml.contains("<m:den>"));
        }
        assert_eq!(
            reopened
                .layout()
                .expect("reopened layout")
                .layout
                .pages
                .len(),
            1
        );
        assert!(reopened.to_pdf().expect("PDF").starts_with(b"%PDF"));
    }

    #[test]
    fn invalid_or_lossy_equations_fail_export_instead_of_becoming_plain_text() {
        for (tex, mathml) in [
            (r"\frac{1}{", None),
            (r"\unknowncommand{x}", None),
            (r"\boxed{x}", Some(r#"<math xmlns="http://www.w3.org/1998/Math/MathML"><menclose notation="box"><mi>x</mi></menclose></math>"#)),
            (r"\\frac{5}{6}", Some(r#"<math xmlns="http://www.w3.org/1998/Math/MathML"><mspace linebreak="newline"/><mi>f</mi><mi>r</mi><mi>a</mi><mi>c</mi><mn>56</mn></math>"#)),
        ] {
            let mut inline = serde_json::json!({ "kind": "math", "tex": tex });
            if let Some(mathml) = mathml { inline["mathml"] = mathml.into(); }
            let ir = parse(&serde_json::json!({
                "version": 1, "locale": "en", "kind": "test", "title": "Invalid math",
                "blocks": [{ "kind": "paragraph", "content": [inline] }],
            }).to_string());
            let Err(error) = render(&ir) else { panic!("silently exported {tex}"); };
            assert!(error.contains("invalid or lossy math"), "{error}");
        }
    }

    #[test]
    fn short_tasks_are_chained_with_keep_with_next_but_release_the_last_line() {
        let (doc, _) = render(&sample_test()).expect("render");
        let keeps: Vec<(String, bool)> = doc.paragraphs().iter().map(|p| (p.text(), p.keep_with_next_value().unwrap_or(false))).collect();
        let prompt2 = keeps.iter().position(|(t, _)| t.starts_with("2.")).expect("task 2 prompt");
        assert!(keeps[prompt2].1, "prompt keeps with its answer lines");
        // The three ruled lines chain; the trailing gap paragraph releases the
        // chain so the following task is not dragged along.
        let lines: Vec<bool> = keeps[prompt2 + 1..prompt2 + 5].iter().map(|(_, k)| *k).collect();
        assert_eq!(lines, vec![true, true, true, false], "{:?}", &keeps[prompt2..prompt2 + 5]);
        // Task 1 (prompt + options table) also keeps its prompt with the table.
        let prompt1 = keeps.iter().position(|(t, _)| t.starts_with("1.")).expect("task 1 prompt");
        assert!(keeps[prompt1].1);
    }

    /// 1×1 transparent PNG.
    const PIXEL_PNG_B64: &str = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWNgYGD4DwABBAEAHnOcQAAAAABJRU5ErkJggg==";

    #[test]
    fn a_watermark_is_drawn_at_the_footer_s_right_end_on_every_page() {
        let mut ir = sample_test();
        ir.watermark = Some(serde_json::from_str(&format!(
            r#"{{ "image": {{ "data": "{PIXEL_PNG_B64}", "mime": "image/png", "alt": "Brand" }}, "widthMm": 20 }}"#
        )).unwrap());
        let (doc, report) = render(&ir).expect("render");
        assert!(report.warnings.is_empty(), "warnings: {:?}", report.warnings);

        let page = doc.layout_page(0).expect("layout").expect("one page");
        let mut footer_images = Vec::new();
        let mut page_number_x = None;
        oxml_layout::walk(&page.elements, &mut |element, transform| match element {
            oxml_layout::PositionedElement::Image { rect, .. } if transform.f + rect.y > page.height * 0.85 => {
                footer_images.push((transform.e + rect.x, rect.width));
            }
            oxml_layout::PositionedElement::Text(run) if transform.f + run.origin.y > page.height * 0.85 && run.field_kind.is_some() => {
                page_number_x = Some(transform.e + run.origin.x);
            }
            _ => {}
        });
        let [(mark_x, mark_w)] = footer_images[..] else { panic!("expected one footer picture, got {footer_images:?}") };
        assert!((mark_w - 20.0 * 72.0 / 25.4).abs() < 0.5, "20 mm wide: {mark_w}");
        let right_margin_x = page.width - t::MARGINS_MM[1] * 72.0 / 25.4;
        assert!((mark_x + mark_w - right_margin_x).abs() < 1.0, "flush with the right margin: {mark_x} + {mark_w} vs {right_margin_x}");
        let page_number_x = page_number_x.expect("page number field in the footer");
        assert!(page_number_x < mark_x - FOOTER_MARK_GAP_MM * 72.0 / 25.4 * 0.9, "page number sits left of the mark: {page_number_x} vs {mark_x}");
    }

    #[test]
    fn pictures_load_from_a_path_and_need_exactly_one_source() {
        let dir = std::env::temp_dir().join(format!("lyset-path-image-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let file = dir.join("pixel.png");
        std::fs::write(&file, base64::engine::general_purpose::STANDARD.decode(PIXEL_PNG_B64).unwrap()).unwrap();

        let ir = parse(&format!(
            r##"{{ "version": 1, "locale": "cs", "kind": "lesson", "title": "Obrázky", "blocks": [
                {{ "kind": "image", "image": {{ "path": {path}, "mime": "image/png" }} }},
                {{ "kind": "image", "image": {{ "data": "{PIXEL_PNG_B64}", "path": {path}, "mime": "image/png" }} }},
                {{ "kind": "image", "image": {{ "mime": "image/png" }} }},
                {{ "kind": "image", "image": {{ "path": {missing}, "mime": "image/png" }} }}
            ] }}"##,
            path = serde_json::to_string(&file).unwrap(),
            missing = serde_json::to_string(&dir.join("missing.png")).unwrap(),
        ));
        let (doc, report) = render(&ir).expect("render");
        assert_eq!(doc.images().len(), 1, "only the well-formed path picture is embedded");
        assert_eq!(report.warnings.len(), 3, "{:?}", report.warnings);
        assert!(report.warnings.iter().any(|w| w.contains("both data and path")));
        assert!(report.warnings.iter().any(|w| w.contains("neither data nor path")));
        assert!(report.warnings.iter().any(|w| w.contains("missing.png")));
        std::fs::remove_dir_all(&dir).unwrap();
    }

    #[test]
    fn a_broken_watermark_image_warns_but_never_blocks_the_document() {
        let mut ir = sample_test();
        ir.watermark = Some(serde_json::from_str(r#"{ "image": { "data": "not-base64!", "mime": "image/png" } }"#).unwrap());
        let (doc, report) = render(&ir).expect("render");
        assert_eq!(report.warnings.len(), 1, "{:?}", report.warnings);
        assert!(report.warnings[0].starts_with("watermark:"));
        assert_eq!(doc.layout().expect("layout").layout.pages.len(), 1);
    }

    #[test]
    fn unknown_ir_fields_are_rejected() {
        let err = serde_json::from_str::<PrintDocument>(r#"{"version":1,"locale":"cs","kind":"test","title":"x","blocks":[],"bogus":1}"#).unwrap_err();
        assert!(err.to_string().contains("bogus"));
    }

    #[test]
    fn a_long_lesson_paginates_with_repeating_footer_fields() {
        let mut blocks = String::new();
        for i in 0..40 {
            blocks.push_str(&format!(
                r#"{{ "kind": "heading", "level": 3, "content": [{{ "kind": "text", "text": "Krok {i}" }}] }},
                    {{ "kind": "paragraph", "content": [{{ "kind": "text", "text": "{}" }}] }},"#,
                "Lorem ipsum dolor sit amet, consectetur adipiscing elit. ".repeat(4)
            ));
        }
        blocks.pop();
        let ir = parse(&format!(
            r##"{{ "version": 1, "locale": "en", "kind": "lesson", "title": "Long",
                 "footer": {{ "left": "Long", "pageNumbers": true, "pageNumberFormat": "Page {{page}} / {{pages}}" }},
                 "blocks": [ {{ "kind": "heading", "level": 1, "content": [{{ "kind": "text", "text": "Long" }}], "kicker": "Lesson plan" }}, {blocks} ] }}"##
        ));
        let (doc, _) = render(&ir).expect("render");
        let layout = doc.layout().expect("layout");
        assert!(layout.layout.pages.len() >= 3, "expected several pages, got {}", layout.layout.pages.len());
        // Every step heading keeps with its paragraph so no heading is orphaned at a page bottom.
        for p in doc.paragraphs() {
            if p.text().starts_with("Krok ") {
                assert_eq!(p.keep_with_next_value(), Some(true), "heading {:?}", p.text());
            }
        }
    }
}
