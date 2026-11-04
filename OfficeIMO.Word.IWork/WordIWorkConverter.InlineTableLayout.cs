using System.Threading;
using OfficeIMO.IWork;

namespace OfficeIMO.Word.IWork;

public static partial class WordIWorkConverter {
    private static HashSet<ulong> FindStandaloneInlineTableBlocks(IWorkPagesProjection projection,
        out HashSet<ulong> inlineDrawables, CancellationToken cancellationToken) {
        inlineDrawables = new HashSet<ulong>();
        var tables = new HashSet<ulong>();
        foreach (IWorkTextParagraph paragraph in projection.Body.Paragraphs) {
            cancellationToken.ThrowIfCancellationRequested();
            IWorkInlineObject? attachment = null;
            int attachmentCount = 0;
            foreach (IWorkTextRun run in paragraph.Runs) {
                cancellationToken.ThrowIfCancellationRequested();
                if (run.InlineObject == null) continue;
                attachment = run.InlineObject;
                attachmentCount++;
                inlineDrawables.Add(attachment.Drawable.RecordIdentifier);
            }
            IWorkParagraphStyle style = paragraph.Style;
            if (attachmentCount == 1 && attachment!.Drawable.MessageType == 6000
                && paragraph.Text.Length == 0 && paragraph.ListLevel < 0
                && style.Alignment is null or IWorkTextAlignment.Natural or IWorkTextAlignment.Left
                && IsZeroOrUnspecified(style.FirstLineIndentPoints)
                && IsZeroOrUnspecified(style.LeftIndentPoints) && IsZeroOrUnspecified(style.RightIndentPoints)
                && IsZeroOrUnspecified(style.SpaceBeforePoints) && IsZeroOrUnspecified(style.SpaceAfterPoints)
                && style.LineSpacingMultiplier is null or 1d
                && style.PageBreakBefore != true && style.KeepWithNext != true && style.KeepLinesTogether != true) {
                tables.Add(attachment.Drawable.RecordIdentifier);
            }
        }
        // Floating objects still use body paragraphs as their page anchors. Keep
        // those hosts until the floating-placement owner can anchor to table blocks.
        foreach (IWorkPagesDrawable drawable in projection.Drawables) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!inlineDrawables.Contains(DrawableIdentifier(drawable))) {
                tables.Clear();
                break;
            }
        }
        return tables;
    }

    private static bool IsZeroOrUnspecified(double? value) => value is null or 0d;

    private static bool HasTableGeometry(IWorkTable table) => table.Geometry is { } geometry
        && (Math.Abs(geometry.LeftPoints) > 0.000001d || Math.Abs(geometry.TopPoints) > 0.000001d
            || Math.Abs(geometry.WidthPoints) > 0.000001d || Math.Abs(geometry.HeightPoints) > 0.000001d
            || Math.Abs(geometry.RotationDegrees) > 0.000001d);

    private static bool CanApplyInlineTableGeometry(IWorkTable table, IWorkPageLayout? pageLayout,
        HashSet<ulong> standaloneTables, CancellationToken cancellationToken) {
        if (table.SourceIdentity == null || !standaloneTables.Contains(table.SourceIdentity.RecordIdentifier)
            || table.Geometry is not { } geometry || pageLayout == null || !CanApplyPageLayout(pageLayout, true)
            || table.AutoResizeRows != false || table.MergedRanges.Count != 0
            || Math.Abs(geometry.LeftPoints) > 0.000001d || Math.Abs(geometry.TopPoints) > 0.000001d
            || Math.Abs(geometry.RotationDegrees) > 0.000001d
            || geometry.WidthPoints <= 0 || geometry.HeightPoints <= 0) return false;
        double width = 0, height = 0;
        long serializedWidth = 0;
        for (int column = 1; column <= table.ColumnCount; column++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (table.GetColumnWidth(column) is not double size || size <= 0 || !FitsSignedTwips(size, true)) return false;
            width += size;
            serializedWidth += ToSignedTwips(size);
        }
        for (int row = 1; row <= table.RowCount; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (table.GetRowHeight(row) is not double size || size <= 0 || !FitsSignedTwips(size, true)) return false;
            height += size;
        }
        return serializedWidth is > 0 and <= int.MaxValue
            && Math.Abs(width - geometry.WidthPoints) <= 0.000001d
            && Math.Abs(height - geometry.HeightPoints) <= 0.000001d
            && width <= pageLayout.WidthPoints - pageLayout.LeftMarginPoints - pageLayout.RightMarginPoints
            && height <= pageLayout.HeightPoints - pageLayout.TopMarginPoints - pageLayout.BottomMarginPoints;
    }

    private static void ApplyInlineTableGrid(WordTable target, IWorkTable source, CancellationToken cancellationToken) {
        var widths = new List<int>(source.ColumnCount);
        for (int column = 1; column <= source.ColumnCount; column++) {
            cancellationToken.ThrowIfCancellationRequested();
            widths.Add(ToSignedTwips(source.GetColumnWidth(column)!.Value));
        }
        target.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus((long)widths.Sum() * 635));
        target.LayoutMode = WordTableLayoutMode.Fixed;
        target.GridColumnWidth = widths.Select(width => global::OfficeIMO.OfficeLength.Emus((long)width * 635)).ToList();
    }
}
