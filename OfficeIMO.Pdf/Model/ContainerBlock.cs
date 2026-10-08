namespace OfficeIMO.Pdf;

internal sealed class ContainerBlock : IPdfBlock {
    public ContainerBlock(IEnumerable<IPdfBlock> blocks, PdfPanelStyle? style, bool useDefaultPanelStyle = false) {
        Guard.NotNull(blocks, nameof(blocks));
        Blocks = blocks.ToList().AsReadOnly();
        Style = style?.Clone() ?? new PdfPanelStyle();
        UseDefaultPanelStyle = useDefaultPanelStyle && style == null;
    }

    public IReadOnlyList<IPdfBlock> Blocks { get; }
    public PdfPanelStyle Style { get; }
    public bool UseDefaultPanelStyle { get; }

    internal TableBlock? FrameTable { get; }
    internal double FrameTableIndent { get; }
    internal double FrameTableHorizontalOffset { get; }
    internal PdfCellBorder? FrameTableBorder { get; }
    internal double FrameTableBorderInset { get; }
    internal double FrameTableContinuationBottomPadding { get; }

    /// <summary>Uses ordinary container pagination while sizing the perimeter from the enclosed grid.</summary>
    internal ContainerBlock(TableBlock table) {
        PdfTableStyle source = table.Style!;
        PdfTableBorderFrame frame = source.BorderFrame!;
        FrameTableIndent = source.LeftIndent;
        FrameTableHorizontalOffset = source.HorizontalOffset;
        FrameTableBorder = frame.Border?.Clone();
        FrameTableBorderInset = frame.HorizontalInset;
        FrameTableContinuationBottomPadding = frame.Spacing / 2D +
            (frame.Border?.Bottom == true ? frame.Border.BottomBorderSnapshot?.PaintThickness ?? 0D : 0D);
        bool repeatsHeader = (source.RepeatHeaderRowCount ?? source.HeaderRowCount) > 0 &&
            table.Rows.Count > source.HeaderRowCount;
        double topBorderThickness = frame.Border?.Top == true ? frame.Border.TopBorderSnapshot?.PaintThickness ?? 0D : 0D;
        Style = new PdfPanelStyle {
            Background = frame.Background,
            PaddingX = frame.Spacing + frame.HorizontalInset, PaddingY = frame.Spacing,
            PaddingTopOverride = frame.Spacing + topBorderThickness,
            // A continued body row starts at half the gap; repeated headers
            // preserve the first fragment's complete leading spacing.
            ContinuationPaddingTopOverride = (repeatsHeader ? frame.Spacing : frame.Spacing / 2D) + topBorderThickness,
            PaddingBottomOverride = frame.Spacing + (frame.Border?.Bottom == true ? frame.Border.BottomBorderSnapshot?.PaintThickness ?? 0D : 0D),
            FragmentBottomInset = frame.Spacing + (frame.Border?.Bottom == true ? frame.Border.BottomBorderSnapshot?.PaintThickness ?? 0D : 0D),
            SpacingBefore = source.SpacingBefore, SpacingAfter = source.SpacingAfter,
            KeepTogether = source.KeepTogether, KeepWithNext = source.KeepWithNext
        };
        PdfTableStyle innerStyle = source.Clone();
        innerStyle.BorderFrame = null;
        innerStyle.LeftIndent = 0D;
        innerStyle.HorizontalOffset = 0D;
        innerStyle.SpacingBefore = innerStyle.SpacingAfter = 0D;
        innerStyle.KeepTogether = innerStyle.KeepWithNext = false;
        FrameTable = new TableBlock(table.Cells.Select(row => row.ToArray()), table.Align, innerStyle);
        foreach (var link in table.Links) FrameTable.AddLink(link.Key, link.Value);
        Blocks = new[] { FrameTable };
    }
}
