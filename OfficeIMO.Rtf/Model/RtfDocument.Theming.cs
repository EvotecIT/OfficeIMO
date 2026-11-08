using OfficeIMO.Drawing;

namespace OfficeIMO.Rtf;

public sealed partial class RtfDocument {
    private Dictionary<RtfRun, AppliedThemeRun>? _appliedThemeRuns;
    private Dictionary<RtfStyle, int>? _appliedThemeStyleFonts;
    private int? _appliedThemeDefaultFont;

    /// <summary>Explicitly applies shared theme typography, paragraph spacing and table appearance to this document.</summary>
    /// <remarks>
    /// Preserves content, emphasis, links, fields, media, numbering, declared headers and page/table geometry.
    /// Monospace and symbol faces are retained. Heading roles come from outline levels, never their appearance.
    /// Theme-owned fonts and base bold settings can be replaced by a later call on this model; after reloading,
    /// loaded formatting is treated as authored. Loading, saving and converting do not call this method automatically.
    /// </remarks>
    /// <returns>The current document.</returns>
    public RtfDocument ApplyTheme(OfficeTheme theme) {
        ArgumentNullException.ThrowIfNull(theme);
        RtfTableTraversalGuard.ValidateDocument(this);
        var application = new ThemeApplication(this, theme, _appliedThemeRuns, _appliedThemeStyleFonts,
            _appliedThemeDefaultFont.HasValue && Settings.DefaultFontId == _appliedThemeDefaultFont);
        application.Apply();
        _appliedThemeRuns = application.AppliedRuns;
        _appliedThemeStyleFonts = application.AppliedStyleFonts;
        _appliedThemeDefaultFont = Settings.DefaultFontId;
        return this;
    }

    private readonly struct AppliedThemeRun {
        internal AppliedThemeRun(int? fontId, bool? bold, bool? originalBold) {
            FontId = fontId;
            Bold = bold;
            OriginalBold = originalBold;
        }

        internal int? FontId { get; }
        internal bool? Bold { get; }
        internal bool? OriginalBold { get; }
    }

    private sealed class ThemeApplication {
        private readonly RtfDocument _document;
        private readonly OfficeTheme _theme;
        private readonly Dictionary<RtfRun, AppliedThemeRun>? _previous;
        private readonly Dictionary<RtfStyle, int>? _previousStyleFonts;
        private readonly HashSet<int> _usedParagraphStyles = new HashSet<int> { 0 };
        private readonly HashSet<object> _visited = new HashSet<object>();
        private readonly int? _originalDefaultFont;
        private readonly bool _ownsDefaultFont;
        private readonly int _bodyFont, _headingFont, _bodyBefore, _bodyAfter, _lineSpacing, _padding, _borderWidth;

        internal ThemeApplication(RtfDocument document, OfficeTheme theme, Dictionary<RtfRun, AppliedThemeRun>? previous,
            Dictionary<RtfStyle, int>? previousStyleFonts, bool ownsDefaultFont) {
            _document = document;
            _theme = theme;
            _previous = previous;
            _previousStyleFonts = previousStyleFonts;
            _ownsDefaultFont = ownsDefaultFont;
            _originalDefaultFont = document.Settings.DefaultFontId ?? document.Fonts.FirstOrDefault()?.Id;
            _bodyBefore = Twips(theme.Spacing.ParagraphBefore);
            _bodyAfter = Twips(theme.Spacing.ParagraphAfter);
            _lineSpacing = checked((int)Math.Round(theme.TextStyle.LineSpacing * 240D, MidpointRounding.AwayFromZero));
            _padding = Twips(theme.TableStyle.CellPadding);
            _borderWidth = Twips(theme.TableStyle.BorderWidth);
            _bodyFont = document.AddFont(theme.Fonts.Body.Family);
            _headingFont = document.AddFont(theme.Fonts.Heading.Family);
        }

        internal Dictionary<RtfRun, AppliedThemeRun> AppliedRuns { get; } = new Dictionary<RtfRun, AppliedThemeRun>();
        internal Dictionary<RtfStyle, int> AppliedStyleFonts { get; } = new Dictionary<RtfStyle, int>();

        internal void Apply() {
            ApplyBlocks(_document.Blocks, null);
            foreach (RtfHeaderFooter story in _document.HeaderFooters)
                foreach (RtfParagraph paragraph in story.Paragraphs) ApplyParagraph(paragraph, null, null, false);
            foreach (RtfNote note in _document.Notes) ApplyNote(note);
            ApplyParagraphStyles();
            _document.Settings.DefaultFontId = _bodyFont;
        }

        private void ApplyBlocks(IReadOnlyList<IRtfBlock> blocks, bool? tableHeader) {
            foreach (IRtfBlock block in blocks) {
                if (block is RtfParagraph paragraph) ApplyParagraph(paragraph, null, tableHeader, false);
                else if (block is RtfTable table) ApplyTable(table);
                else if (block is RtfShape shape) ApplyShape(shape);
                else if (block is RtfObject embedded) ApplyParagraph(embedded.Result, null, tableHeader, false);
            }
        }

        private void ApplyParagraph(RtfParagraph paragraph, RtfParagraph? containingParagraph, bool? tableHeader, bool hyperlink) {
            if (!_visited.Add(paragraph)) return;
            RtfParagraph context = containingParagraph ?? paragraph;
            foreach (RtfStyle style in _document.GetFormattingStyleChain(context.StyleId ?? 0, RtfStyleKind.Paragraph))
                _usedParagraphStyles.Add(style.Id);
            int? outline = _document.GetParagraphFormatting(context).OutlineLevel;
            int heading = outline is >= 0 and <= 8 ? Math.Min(outline.Value + 1, 6) : 0;
            if (containingParagraph == null) {
                OfficeThemeParagraphSpacing? spacing = heading > 0 ? _theme.TextStyle.GetHeadingSpacing(heading) : null;
                paragraph.SpaceBeforeTwips = tableHeader.HasValue ? 0 : spacing != null ? Twips(spacing.Before) : _bodyBefore;
                paragraph.SpaceAfterTwips = tableHeader.HasValue ? 0 : spacing != null ? Twips(spacing.After) : _bodyAfter;
                paragraph.SpaceBeforeAuto = false;
                paragraph.SpaceAfterAuto = false;
                paragraph.LineSpacingTwips = tableHeader.HasValue ? 240 : _lineSpacing;
                paragraph.LineSpacingMultiple = true;
            }

            foreach (IRtfInline inline in paragraph.Inlines) {
                if (inline is RtfRun run) {
                    ApplyRun(context, run, heading, tableHeader, hyperlink || run.Hyperlink != null);
                    if (run.Note != null) ApplyNote(run.Note);
                } else if (inline is RtfField field) {
                    ApplyParagraph(field.Result, context, tableHeader, hyperlink || field.HyperlinkField != null);
                } else if (inline is RtfObject embedded) {
                    ApplyParagraph(embedded.Result, context, tableHeader, hyperlink);
                } else if (inline is RtfShape shape) ApplyShape(shape);
                else if (inline is RtfGeneratedText generated && generated.Note != null) ApplyNote(generated.Note);
            }
            // ListText is a marker payload; its face and geometry belong to the numbering definition.
        }

        private void ApplyRun(RtfParagraph context, RtfRun run, int heading, bool? tableHeader, bool hyperlink) {
            AppliedThemeRun previous = default;
            bool hadPrevious = _previous != null && _previous.TryGetValue(run, out previous);
            if (hadPrevious && previous.Bold.HasValue && run.DirectBold == previous.Bold)
                run.DirectBold = previous.OriginalBold;
            RtfRun effective = _document.GetRunFormatting(context, run);
            int? fontId = effective.FontId ?? _originalDefaultFont;
            RtfFont? font = _document.Fonts.FirstOrDefault(item => item.Id == fontId);
            bool ownsFont = hadPrevious && previous.FontId.HasValue && run.FontId == previous.FontId || OwnsInheritedFont(context, run);
            bool semanticFont = PreservesFace(font);
            int? appliedFont = null;
            if (semanticFont && !ownsFont) run.FontId = fontId;
            else run.FontId = appliedFont = !tableHeader.HasValue && heading > 0 ? _headingFont : _bodyFont;
            run.FontSize = tableHeader.HasValue ? Math.Max(1D, _theme.TypeScale.Body.ToPoints() - 1D) : HeadingSize(heading);
            run.ForegroundColorIndex = Color(hyperlink ? _theme.TextStyle.LinkColor : tableHeader == true
                ? _theme.TableStyle.HeaderText : heading > 0 ? PageColor(_theme.TextStyle.HeadingColor) : PageText());

            bool? baseBold = tableHeader == true ? _theme.TableStyle.HeaderBold : !tableHeader.HasValue && heading > 0
                ? _theme.TextStyle.HeadingBold : (bool?)null;
            bool? originalBold = run.DirectBold;
            bool? appliedBold = null;
            if (baseBold.HasValue && !run.DirectBold.HasValue &&
                !_document.GetFormattingStyleChain(run.StyleId, RtfStyleKind.Character).Any(style => style.Bold.HasValue) &&
                !_document.GetFormattingStyleChain(context.StyleId ?? 0, RtfStyleKind.Paragraph).Any(style => style.Bold.HasValue)) {
                run.DirectBold = appliedBold = baseBold;
            }
            AppliedRuns[run] = new AppliedThemeRun(appliedFont, appliedBold, originalBold);
        }

        private bool OwnsInheritedFont(RtfParagraph paragraph, RtfRun run) {
            if (run.FontId.HasValue) return false;
            if (run.UseDefaultCharacterFormatting) return _ownsDefaultFont;
            if (_document.GetFormattingStyleChain(run.StyleId, RtfStyleKind.Character).Any(style => style.FontId.HasValue)) return false;
            RtfStyle? owner = _document.GetFormattingStyleChain(paragraph.StyleId ?? 0, RtfStyleKind.Paragraph)
                .Reverse().FirstOrDefault(style => style.FontId.HasValue);
            return owner == null ? _ownsDefaultFont : _previousStyleFonts != null &&
                _previousStyleFonts.TryGetValue(owner, out int assigned) && owner.FontId == assigned;
        }

        private double HeadingSize(int level) => Math.Max(1D, (level switch {
            1 => _theme.TypeScale.Heading1,
            2 => _theme.TypeScale.Heading2,
            3 => _theme.TypeScale.Heading3,
            4 => _theme.TypeScale.Heading4,
            5 => _theme.TypeScale.Heading5,
            6 => _theme.TypeScale.Heading6,
            _ => _theme.TypeScale.Body
        }).ToPoints());

        private OfficeColor PageText() => _theme.IsDark ? _theme.Palette.Background : _theme.Palette.Text;
        private OfficeColor PageColor(OfficeColor color) => _theme.IsDark && color.Role == OfficeColorRole.Text ? PageText() : color;

        private static bool PreservesFace(RtfFont? font) => font != null && (font.Charset == 2 || font.Pitch == 1 ||
            font.Family is RtfFontFamily.Modern or RtfFontFamily.Technical ||
            OfficeDocumentFontSubstitutions.PreservesSemanticFont(font.Name));

        private void ApplyParagraphStyles() {
            if (!_document.Styles.Any(style => style.Id == 0 && style.Kind == RtfStyleKind.Paragraph))
                _document.AddStyle(0, "Normal");
            // Snapshot effective faces and their ownership before changing any base style.
            // Generated controls and blank lines depend on this catalog rather than a frozen text run.
            var originals = _document.Styles.Where(style => style.Kind == RtfStyleKind.Paragraph && _usedParagraphStyles.Contains(style.Id))
                .Select(style => {
                    var paragraph = new RtfParagraph { StyleId = style.Id };
                    var run = new RtfRun(string.Empty);
                    return new {
                        Style = style,
                        FontId = _document.GetRunFormatting(paragraph, run).FontId ?? _originalDefaultFont,
                        OwnsFont = OwnsInheritedFont(paragraph, run)
                    };
                }).ToArray();
            foreach (var original in originals) {
                RtfStyle style = original.Style;
                int? outline = _document.GetParagraphFormatting(new RtfParagraph { StyleId = style.Id }).OutlineLevel;
                int heading = outline is >= 0 and <= 8 ? Math.Min(outline.Value + 1, 6) : 0;
                if (!original.OwnsFont && PreservesFace(_document.Fonts.FirstOrDefault(font => font.Id == original.FontId)))
                    style.FontId = original.FontId;
                else {
                    style.FontId = heading > 0 ? _headingFont : _bodyFont;
                    AppliedStyleFonts[style] = style.FontId.Value;
                }
                style.FontSize = HeadingSize(heading);
                style.ForegroundColorIndex = Color(heading > 0 ? PageColor(_theme.TextStyle.HeadingColor) : PageText());
            }
        }

        private void ApplyNote(RtfNote note) {
            foreach (RtfParagraph paragraph in note.Paragraphs) ApplyParagraph(paragraph, null, null, false);
        }

        private void ApplyShape(RtfShape shape) {
            foreach (RtfParagraph paragraph in shape.TextBoxParagraphs) ApplyParagraph(paragraph, null, null, false);
        }

        private int Color(OfficeColor value) {
            OfficeColor color = value.Resolve(_theme);
            int index = _document._colors.FindIndex(item => !item.IsAutomatic && item.ThemeColor == null &&
                item.Tint == null && item.Shade == null && item.Red == color.R && item.Green == color.G && item.Blue == color.B);
            return index >= 0 ? index + 1 : _document.AddColor(color.R, color.G, color.B);
        }

        private static int Twips(OfficeLength value) => checked((int)value.ToTwips());

        private void ApplyTable(RtfTable table) {
            if (!_visited.Add(table)) return;
            OfficeThemeTableStyle style = _theme.TableStyle;
            int bodyIndex = 0;
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                RtfTableRow row = table.Rows[rowIndex];
                bool header = row.RepeatHeader;
                bool stripe = !header && style.Banding && bodyIndex++ % 2 == 1;
                OfficeColor fill = (header ? style.HeaderFill : stripe ? _theme.IsDark
                    ? OfficeTheme.Modern.TableStyle.BandFill : style.BandFill : OfficeColor.Transparent).Resolve(_theme);
                row.BackgroundColorIndex = null;
                row.ShadingForegroundColorIndex = null;
                row.ShadingPatternValue = null;
                row.ShadingPatternPercent = null;
                row.ShadingPattern = RtfShadingPattern.None;
                row.PaddingTopTwips = row.PaddingBottomTwips = row.PaddingLeftTwips = row.PaddingRightTwips = _padding;
                ClearBorder(row.TopBorder); ClearBorder(row.LeftBorder); ClearBorder(row.BottomBorder);
                ClearBorder(row.RightBorder); ClearBorder(row.HorizontalBorder); ClearBorder(row.VerticalBorder);
                bool headerEnd = header && (rowIndex + 1 == table.Rows.Count || !table.Rows[rowIndex + 1].RepeatHeader);
                foreach (RtfTableCell cell in row.Cells) {
                    cell.BackgroundColorIndex = fill.A == 0 ? null : Color(fill);
                    cell.ShadingForegroundColorIndex = null;
                    cell.ShadingPatternPercent = null;
                    cell.ShadingPattern = RtfShadingPattern.None;
                    cell.PaddingTopTwips = cell.PaddingBottomTwips = cell.PaddingLeftTwips = cell.PaddingRightTwips = _padding;
                    bool grid = style.BorderMode == OfficeThemeTableBorderMode.Grid;
                    SetBorder(cell.TopBorder, grid, _borderWidth, style.BorderColor);
                    SetBorder(cell.LeftBorder, grid, _borderWidth, style.BorderColor);
                    SetBorder(cell.RightBorder, grid, _borderWidth, style.BorderColor);
                    SetBorder(cell.BottomBorder, style.BorderMode != OfficeThemeTableBorderMode.None, _borderWidth, style.BorderColor);
                    if (headerEnd && style.HeaderRule.IsVisible)
                        SetBorder(cell.BottomBorder, true, Twips(style.HeaderRule.Width), style.HeaderRule.Color);
                    SetBorder(cell.TopLeftToBottomRightBorder, false, 0, style.BorderColor);
                    SetBorder(cell.TopRightToBottomLeftBorder, false, 0, style.BorderColor);
                    ApplyBlocks(cell.Blocks, header);
                }
            }
        }

        private void SetBorder(RtfTableCellBorder border, bool visible, int width, OfficeColor color) {
            border.Style = visible && width > 0 ? RtfTableCellBorderStyle.Single : RtfTableCellBorderStyle.None;
            border.Width = visible ? width : 0;
            border.ColorIndex = visible ? Color(color) : (int?)null;
        }

        private static void ClearBorder(RtfTableRowBorder border) {
            border.Style = RtfTableCellBorderStyle.None;
            border.Width = 0;
            border.ColorIndex = null;
        }
    }
}
