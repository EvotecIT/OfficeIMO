namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    internal PdfFootnoteDefinition DefineFootnote(Action<PdfContentBuilder> build, PdfFootnoteOptions? options) {
        Guard.NotNull(build, nameof(build));
        options ??= new PdfFootnoteOptions();
        if (options.Placement is not PdfFootnotePlacement.BottomOfPage and not PdfFootnotePlacement.BeneathText)
            throw new ArgumentOutOfRangeException(nameof(options), "Unknown footnote placement.");
        // Capture the option values before callbacks run. Retained builders are detached below.
        var snapshot = new PdfFootnoteOptions { Placement = options.Placement };
        Action<PdfContentBuilder>? separator = options.Separator;
        Action<PdfContentBuilder>? continuationSeparator = options.ContinuationSeparator;
        var content = BuildFlowBlocks(build);
        IReadOnlyList<IPdfBlock> first = separator == null ? Array.Empty<IPdfBlock>() : BuildFlowBlocks(separator);
        IReadOnlyList<IPdfBlock> continuation = continuationSeparator == null ? Array.Empty<IPdfBlock>() : BuildFlowBlocks(continuationSeparator);
        ValidateFootnoteBlocks(content, _options);
        ValidateFootnoteBlocks(first, _options);
        ValidateFootnoteBlocks(continuation, _options);
        if (content.Count == 0) throw new ArgumentException("A footnote must contain flow content.", nameof(build));
        return new PdfFootnoteDefinition(this, content, snapshot, first, continuation);
    }

    internal static void ValidateFootnoteBlocks(IReadOnlyList<IPdfBlock> blocks, PdfOptions options) {
        foreach (IPdfBlock block in blocks) {
            switch (block) {
                case RichParagraphBlock paragraph:
                    PdfParagraphStyle? paragraphStyle = paragraph.Style ?? options.DefaultParagraphStyleSnapshot;
                    if (paragraphStyle?.AnchoredCanvas != null || paragraph.Runs.Any(run => run.FootnoteReference != null || run.FormField != null))
                        throw new NotSupportedException("Footnotes cannot contain recursive references, form fields, or anchored canvas content.");
                    if (paragraphStyle?.KeepWithNext == true)
                        throw new NotSupportedException("Footnote paragraphs do not support keep-with-next chains. Use paragraph keep-together or widow/orphan constraints.");
                    break;
                case HorizontalRuleBlock rule:
                    if ((rule.Style ?? options.DefaultHorizontalRuleStyleSnapshot)?.KeepWithNext == true)
                        throw new NotSupportedException("Footnote horizontal rules do not support keep-with-next chains.");
                    break;
                case ImageBlock image:
                    if ((image.Style ?? options.DefaultImageStyleSnapshot)?.KeepWithNext == true)
                        throw new NotSupportedException("Footnote images do not support keep-with-next chains.");
                    break;
                case SpacerBlock:
                case BookmarkBlock:
                    break;
                case FlowBlock { IsReplayable: false, StaticBlocks: { } children }:
                    var flow = (FlowBlock)block;
                    if (flow.Options.KeepTogether || flow.Options.ShowIf != null || flow.Options.MinimumRemainingHeight > 0D || flow.Options.OverflowBehavior != PdfFlowOverflowBehavior.Continue)
                        throw new NotSupportedException("Footnote flow groups cannot introduce conditional or whole-group pagination constraints.");
                    ValidateFootnoteBlocks(children, options);
                    break;
                default:
                    throw new NotSupportedException("Footnote flow does not support " + block.GetType().Name + ". Use paragraphs, images, rules, and spacers.");
            }
        }
    }

    private void ValidateFootnoteDocument() {
        if (!_blocks.SelectMany(PdfFootnoteTraversal.References).Any()) return;
        foreach (IPdfBlock block in _blocks) ValidateFootnoteBodyProfile(block, _options);
    }

    internal static void ValidateFootnoteBodyProfile(IPdfBlock block, PdfOptions options) {
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
