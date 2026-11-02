namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    internal PdfFootnoteDefinition DefineFootnote(Action<PdfContentBuilder> build, PdfFootnoteOptions? options) {
        Guard.NotNull(build, nameof(build));
        options ??= new PdfFootnoteOptions();
        if (options.Placement is not PdfFootnotePlacement.BottomOfPage and not PdfFootnotePlacement.BeneathText)
            throw new ArgumentOutOfRangeException(nameof(options), "Unknown footnote placement.");
        // Capture the option values before callbacks run. Retained builders are detached below.
        var snapshot = new PdfFootnoteOptions {
            Placement = options.Placement,
            ReserveContinuationNoticeSpaceOnFinalPage = options.ReserveContinuationNoticeSpaceOnFinalPage
        };
        Action<PdfContentBuilder>? separator = options.Separator;
        Action<PdfContentBuilder>? continuationSeparator = options.ContinuationSeparator;
        Action<PdfContentBuilder>? continuationNotice = options.ContinuationNotice;
        var content = BuildFlowBlocks(build);
        if (options.SharedDecorationOwner != null && !ReferenceEquals(options.SharedDecorationOwner, this))
            throw new ArgumentException("Footnote decoration blocks belong to another document.", nameof(options));
        IReadOnlyList<IPdfBlock> first = options.SharedSeparator ??
            (separator == null ? Array.Empty<IPdfBlock>() : BuildFlowBlocks(separator));
        IReadOnlyList<IPdfBlock> continuation = options.SharedContinuationSeparator ??
            (continuationSeparator == null ? Array.Empty<IPdfBlock>() : BuildFlowBlocks(continuationSeparator));
        IReadOnlyList<IPdfBlock> notice = options.SharedContinuationNotice ??
            (continuationNotice == null ? Array.Empty<IPdfBlock>() : BuildFlowBlocks(continuationNotice));
        ValidateFootnoteBlocks(content, _options);
        ValidateFootnoteBlocks(first, _options, allowTables: false);
        ValidateFootnoteBlocks(continuation, _options, allowTables: false);
        ValidateFootnoteBlocks(notice, _options, allowTables: false);
        if (content.Count == 0) throw new ArgumentException("A footnote must contain flow content.", nameof(build));
        return new PdfFootnoteDefinition(this, content, snapshot, first, continuation, notice);
    }

    internal static void ValidateFootnoteBlocks(IReadOnlyList<IPdfBlock> blocks, PdfCreateOptions options, bool allowTables = true) {
        foreach (IPdfBlock block in blocks) {
            switch (block) {
                case RichParagraphBlock paragraph:
                    PdfParagraphStyle? paragraphStyle = paragraph.Style ?? options.DefaultParagraphStyleSnapshot;
                    if (paragraphStyle?.AnchoredCanvas != null || paragraph.Runs.Any(run => run.FootnoteReference != null || run.FormField != null))
                        throw new NotSupportedException("Footnotes cannot contain recursive references, form fields, or anchored canvas content.");
                    break;
                case HorizontalRuleBlock:
                    break;
                case ImageBlock:
                    break;
                case TableBlock table:
                    if (!allowTables) throw new NotSupportedException("Footnote separator and notice stories cannot contain tables.");
                    ValidateFootnoteTable(table, options);
                    break;
                case PdfListBlock list:
                    if (list.RichItems.SelectMany(item => item.Runs).Any(run => run.FootnoteReference != null || run.FormField != null))
                        throw new NotSupportedException("Footnote lists cannot contain recursive references or form fields.");
                    break;
                case ContainerBlock container:
                    PdfPanelStyle panelStyle = container.UseDefaultPanelStyle ? options.DefaultPanelStyleSnapshot ?? container.Style : container.Style;
                    if (panelStyle.AnchoredCanvas != null)
                        throw new NotSupportedException("Footnote panels require static flow without anchored canvas content.");
                    ValidateFootnoteBlocks(container.Blocks, options, allowTables);
                    break;
                case SpacerBlock:
                case BookmarkBlock:
                    break;
                case FlowBlock { IsReplayable: false, StaticBlocks: { } children }:
                    var flow = (FlowBlock)block;
                    if (flow.Options.KeepTogether || flow.Options.ShowIf != null || flow.Options.MinimumRemainingHeight.ToPoints() > 0D || flow.Options.OverflowBehavior != PdfFlowOverflowBehavior.Continue)
                        throw new NotSupportedException("Footnote flow groups cannot introduce conditional or whole-group pagination constraints.");
                    ValidateFootnoteBlocks(children, options, allowTables);
                    break;
                default:
                    throw new NotSupportedException("Footnote flow does not support " + block.GetType().Name + ". Use paragraphs, eager inline tables, lists, static panels, images, rules, and spacers.");
            }
        }
    }

    /// <summary>Admits eager inline tables whose cell content has a resumable, inert PDF fragment contract.</summary>
    private static void ValidateFootnoteTable(TableBlock table, PdfCreateOptions options) {
        PdfTableStyle style = table.Style ?? options.DefaultTableStyleSnapshot ?? PdfTableStyles.Light();
        if (style.Position != null || !style.ConsumesVerticalFlow || style.Caption != null)
            throw new NotSupportedException("Footnote tables require inline flow without floating placement or captions.");
        for (int row = 0; row < table.Cells.Count; row++) {
            foreach (PdfTableCell cell in table.Cells[row]) {
                if (cell.Blocks.Count > 0) ValidateFootnoteBlocks(cell.Blocks, options);
                if (cell.Runs.Any(run => run.FootnoteReference != null || run.FormField != null) ||
                    cell.Paragraphs.SelectMany(paragraph => paragraph.Runs).Any(run => run.FootnoteReference != null || run.FormField != null) ||
                    cell.FormFields.Count > 0 || cell.CheckBoxes.Count > 0 || cell.Viewport != null || cell.TextRotation != 0)
                    throw new NotSupportedException("Footnote table cells cannot contain recursive references, form fields, viewports, or rotated text.");
                if (cell.RowSpan <= 1) continue;
                if (row < style.HeaderRowCount || row + cell.RowSpan > table.Cells.Count - style.FooterRowCount ||
                    cell.Images.Count > 0 || cell.Runs.Any(run => run.InlineElement is PdfInlineImage) ||
                    cell.Paragraphs.SelectMany(paragraph => paragraph.Runs).Any(run => run.InlineElement is PdfInlineImage) ||
                    style.CellDataBars?.Keys.Any(key => key.Row == row) == true ||
                    style.CellIcons?.Keys.Any(key => key.Row == row) == true ||
                    Enumerable.Range(row, cell.RowSpan).Any(index =>
                        style.FixedRowHeights != null && index < style.FixedRowHeights.Count && style.FixedRowHeights[index].HasValue ||
                        !PdfWriter.GetTableRowAllowBreakAcrossPages(style, index)))
                    throw new NotSupportedException("Footnote row spans require flexible body rows and text-only merged cells.");
            }
        }
    }

    private void ValidateFootnoteDocument() {
        if (!_blocks.SelectMany(PdfFootnoteTraversal.References).Any()) return;
        foreach (IPdfBlock block in _blocks) ValidateFootnoteBodyProfile(block, _options);
    }

    internal static void ValidateFootnoteBodyProfile(IPdfBlock block, PdfCreateOptions options) {
        if (block is MultiColumnBlock or RowBlock or PdfCanvasBlock)
            throw new NotSupportedException("Page-associated footnotes require single-column body flow without page-canvas content.");
        if (block is RichParagraphBlock paragraph && (paragraph.Style ?? options.DefaultParagraphStyleSnapshot)?.AnchoredCanvas != null ||
            block is HeadingBlock { Style.AnchoredCanvas: not null })
            throw new NotSupportedException("Page-associated footnotes do not support anchored canvas body content.");
        if (block is ContainerBlock { Style.AnchoredCanvas: not null })
            throw new NotSupportedException("Page-associated footnotes do not support anchored canvas body content.");
        if (block is TableBlock or DeferredTableBlock) {
            PdfTableStyle? style = (block is TableBlock table ? table.Style : ((DeferredTableBlock)block).Style)
                ?? options.DefaultTableStyleSnapshot;
            if (style?.Position != null || style?.ConsumesVerticalFlow == false)
                throw new NotSupportedException("Page-associated footnotes do not support floating or non-flow body tables.");
        }
        if (block is PageBlock pageOptions) options = pageOptions.Options;
        IEnumerable<IPdfBlock> children = block switch {
            FlowBlock { StaticBlocks: { } content } => content,
            PageBlock page => page.Blocks,
            ContainerBlock container => container.Blocks,
            SemanticBlock semantic => semantic.Blocks,
            LayerBlock layer => layer.Blocks,
            SectionBlock section => section.Blocks,
            _ => Array.Empty<IPdfBlock>()
        };
        foreach (IPdfBlock child in children) ValidateFootnoteBodyProfile(child, options);
    }
}
