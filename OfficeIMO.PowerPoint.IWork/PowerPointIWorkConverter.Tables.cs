using System.Threading;
using OfficeIMO.IWork;

namespace OfficeIMO.PowerPoint.IWork;

public static partial class PowerPointIWorkConverter {
    private static void AddEditableTable(PowerPointSlide slide, IWorkTable source,
        CancellationToken cancellationToken) {
        if (source.RowCount == 0 || source.ColumnCount == 0) return;
        double left = source.Geometry?.LeftPoints ?? 72d;
        double top = source.Geometry?.TopPoints ?? 72d;
        double width = QuantizePositiveEmuPoints(TableAxisExtent(source.ColumnCount,
            source.ColumnWidths, source.DefaultColumnWidth, source.Geometry?.WidthPoints,
            Math.Max(144d, 72d * source.ColumnCount)));
        double height = QuantizePositiveEmuPoints(TableAxisExtent(source.RowCount,
            source.RowHeights, source.DefaultRowHeight, source.Geometry?.HeightPoints,
            Math.Max(36d, 24d * source.RowCount)));
        PowerPointTable table = slide.AddTablePoints(source.RowCount, source.ColumnCount,
            left, top, width, height);
        table.AltText = source.AccessibilityDescription;
        table.Rotation = source.Geometry?.RotationDegrees ?? 0d;
        table.FirstRow = source.HeaderRowCount > 0;
        table.FirstColumn = source.HeaderColumnCount > 0;
        table.LastRow = source.FooterRowCount > 0;
        for (int row = 1; row <= source.RowCount; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int column = 1; column <= source.ColumnCount; column++) {
                PowerPointTableCell fillTarget = table.GetCell(row - 1, column - 1);
                if (source.GetFill(row, column) is { } fill) {
                    if (fill.IsNone) fillTarget.NoFill = true;
                    else fillTarget.FillColor = fill.Color!.RgbHex;
                }
                if (source.GetParagraphStyle(row, column) is { } style) {
                    PowerPointParagraph paragraph = table.GetCell(row - 1, column - 1).Paragraphs[0];
                    ApplyParagraphStyle(paragraph, style, string.Empty);
                    foreach (PowerPointTextRun run in paragraph.Runs) ApplyTextStyle(run, style.TextStyle);
                }
            }
        }
        foreach (IWorkTableCell sourceCell in source.Cells) {
            cancellationToken.ThrowIfCancellationRequested();
            PowerPointTableCell target = table.GetCell(sourceCell.Row - 1, sourceCell.Column - 1);
            if (sourceCell.Padding is { } padding) {
                target.PaddingLeftPoints = padding.LeftPoints;
                target.PaddingTopPoints = padding.TopPoints;
                target.PaddingRightPoints = padding.RightPoints;
                target.PaddingBottomPoints = padding.BottomPoints;
            }
            if (sourceCell.VerticalAlignment is { } vertical) target.VerticalAlignment = vertical switch {
                IWorkCellVerticalAlignment.Top => PowerPointTextVerticalAlignment.Top,
                IWorkCellVerticalAlignment.Middle => PowerPointTextVerticalAlignment.Center,
                _ => PowerPointTextVerticalAlignment.Bottom
            };
            IWorkParagraphStyle? defaultStyle = source.GetParagraphStyle(sourceCell.Row, sourceCell.Column);
            bool header = sourceCell.Row <= source.HeaderRowCount || sourceCell.Column <= source.HeaderColumnCount
                || sourceCell.Row > source.RowCount - source.FooterRowCount;
            if (sourceCell.RichText is { Paragraphs.Count: > 0 } richText) {
                IReadOnlyList<PowerPointParagraph> paragraphs = target.SetParagraphs(
                    richText.Paragraphs.Select(_ => string.Empty));
                var listState = new IWorkPowerPointListState();
                for (int index = 0; index < paragraphs.Count; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    IWorkTextParagraph sourceParagraph = richText.Paragraphs[index];
                    if (defaultStyle != null) ApplyParagraphStyle(paragraphs[index], defaultStyle, sourceParagraph.Text);
                    ApplyParagraphStyle(paragraphs[index], sourceParagraph,
                        listState.StartsAtSourceLabel(sourceParagraph));
                    WriteParagraphContent(paragraphs[index], sourceParagraph, cancellationToken, defaultStyle?.TextStyle, header);
                }
            } else {
                target.Text = sourceCell.Kind == IWorkCellKind.Formula && sourceCell.Value != null
                    ? sourceCell.CachedDisplayText
                    : sourceCell.DisplayText;
                if (header) target.Bold = true;
                if (defaultStyle != null) {
                    foreach (PowerPointParagraph paragraph in target.Paragraphs) {
                        ApplyParagraphStyle(paragraph, defaultStyle, paragraph.Text);
                        foreach (PowerPointTextRun run in paragraph.Runs) ApplyTextStyle(run, defaultStyle.TextStyle);
                    }
                }
            }
        }
        foreach (IWorkTableMergeRange merge in source.MergedRanges) {
            cancellationToken.ThrowIfCancellationRequested();
            table.MergeCells(merge.FirstRow - 1, merge.FirstColumn - 1,
                merge.LastRow - 1, merge.LastColumn - 1);
        }
        ApplyAxisSizing(source.ColumnCount, source.ColumnWidths, source.DefaultColumnWidth,
            source.Geometry?.WidthPoints, width, table.SetColumnWidthPoints, cancellationToken);
        ApplyAxisSizing(source.RowCount, source.RowHeights, source.DefaultRowHeight,
            source.Geometry?.HeightPoints, height, table.SetRowHeightPoints, cancellationToken);
    }

    private static IEnumerable<IWorkParagraphStyle> TableParagraphStyles(IWorkTable table) =>
        new[] { table.TextStyles.Body, table.TextStyles.HeaderRow, table.TextStyles.HeaderColumn, table.TextStyles.FooterRow }
            .Concat(table.Cells.Select(cell => cell.ParagraphStyle)).Where(style => style != null).Cast<IWorkParagraphStyle>();

    private static string? FindTableTextStyleLimitation(IWorkTable table, bool allowPartial) {
        foreach (IWorkParagraphStyle style in TableParagraphStyles(table)) {
            if (!allowPartial && (style.PageBreakBefore == true || style.KeepWithNext == true || style.KeepLinesTogether == true))
                return $"Keynote table '{table.Name}' contains paragraph pagination formatting that the PPTX owner cannot preserve.";
            if (!FitsTextCoordinate(style.FirstLineIndentPoints) || !FitsTextCoordinate(style.LeftIndentPoints)
                || !FitsTextCoordinate(style.RightIndentPoints) || Math.Abs(style.RightIndentPoints.GetValueOrDefault()) > 0.000001d
                || !FitsSpacing(style.SpaceBeforePoints) || !FitsSpacing(style.SpaceAfterPoints))
                return $"Keynote table '{table.Name}' contains paragraph formatting outside the PPTX range.";
            IWorkTextStyle text = style.TextStyle;
            if (text.FontSizePoints is double size && (!IsFinite(size) || size < 1d || size > 4000d
                    || size * 100d != Math.Round(size * 100d, MidpointRounding.AwayFromZero)))
                return $"Keynote table '{table.Name}' contains a font size outside the PPTX range or hundredth-point precision.";
            if (text.Color is { Alpha: < byte.MaxValue } || text.BackgroundColor is { Alpha: < byte.MaxValue })
                return $"Keynote table '{table.Name}' contains transparent text colors that the PPTX owner cannot preserve.";
        }
        return null;
    }

    private static IEnumerable<double> PaddingPoints(IWorkCellPadding padding) {
        yield return padding.LeftPoints;
        yield return padding.TopPoints;
        yield return padding.RightPoints;
        yield return padding.BottomPoints;
    }

    private static double TableAxisExtent(int count, IReadOnlyDictionary<int, double> sizes,
        double? defaultSize, double? geometryExtent, double fallbackExtent) {
        if (geometryExtent is > 0) return geometryExtent.Value;
        if (sizes.Count == 0) return defaultSize is > 0 ? defaultSize.Value * count : fallbackExtent;
        return sizes.Values.Sum() + (count - sizes.Count) * (defaultSize ?? fallbackExtent / count);
    }

    private static double AxisSourceTotal(int count, IReadOnlyDictionary<int, double> sizes,
        double? defaultSize, double targetExtent) =>
        sizes.Values.Sum() + (count - sizes.Count) * (defaultSize ?? targetExtent / count);

    private static void ApplyAxisSizing(int count, IReadOnlyDictionary<int, double> sizes,
        double? defaultSize, double? geometryExtent, double targetExtent,
        Action<int, double> setSize, CancellationToken cancellationToken) {
        // An axis with no explicit sizes retains the existing uniform drawable-extent contract.
        if (sizes.Count == 0 && (geometryExtent is > 0 || defaultSize == null)) return;
        double total = AxisSourceTotal(count, sizes, defaultSize, targetExtent);
        double scale = geometryExtent is > 0 ? targetExtent / total : 1d;
        for (int index = 0; index < count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            double value = sizes.TryGetValue(index + 1, out double size)
                ? size : defaultSize ?? targetExtent / count;
            setSize(index, QuantizePositiveEmuPoints(value * scale));
        }
    }

    private static bool AxisSizingRequiresEmuRounding(int count, IReadOnlyDictionary<int, double> sizes,
        double? defaultSize, double? geometryExtent, double fallbackExtent) {
        if (sizes.Count == 0 && (geometryExtent is > 0 || defaultSize == null)) return false;
        double targetExtent = TableAxisExtent(count, sizes, defaultSize, geometryExtent, fallbackExtent);
        double total = AxisSourceTotal(count, sizes, defaultSize, targetExtent);
        double scale = geometryExtent is > 0 ? targetExtent / total : 1d;
        return sizes.Values.Any(size => !IsExactEmu(size * scale))
            || count > sizes.Count && !IsExactEmu((defaultSize ?? targetExtent / count) * scale);
    }

    private static bool AxisSizingIsScaled(int count, IReadOnlyDictionary<int, double> sizes,
        double? defaultSize, double? geometryExtent) => sizes.Count > 0 && geometryExtent is > 0
        && AxisSourceTotal(count, sizes, defaultSize, geometryExtent.Value) != geometryExtent.Value;
}
