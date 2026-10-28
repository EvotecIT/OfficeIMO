using System.Threading;
using OfficeIMO.IWork;

namespace OfficeIMO.PowerPoint.IWork;

public static partial class PowerPointIWorkConverter {
    private static void AddEditableTable(PowerPointSlide slide, IWorkTable source,
        CancellationToken cancellationToken, Func<string?, string?> resolveHyperlink) {
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
        // Native region styles are expanded below. Destination heading and
        // banding flags would add a different theme over those declarations.
        table.ApplyStyle(PowerPointTableStylePreset.Plain);
        table.StyleId = null;
        var mergedBorders = source.ResolveMergedBorders(cancellationToken).ToDictionary(
            merge => (merge.Range.FirstRow, merge.Range.FirstColumn), merge => merge.Borders);
        for (int row = 1; row <= source.RowCount; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int column = 1; column <= source.ColumnCount; column++) {
                PowerPointTableCell fillTarget = table.GetCell(row - 1, column - 1);
                fillTarget.NoFill = true;
                fillTarget.HideBorders(PowerPointTableCellBorders.All);
                ApplyTableBorders(fillTarget, mergedBorders.TryGetValue((row, column), out var perimeter)
                    ? perimeter : source.GetBorders(row, column));
                if (source.GetFill(row, column) is { } fill) {
                    if (fill.IsNone) fillTarget.NoFill = true;
                    else fillTarget.FillColor = fill.Color!.RgbHex;
                }
                if (source.GetPadding(row, column) is { } padding) {
                    fillTarget.PaddingLeftPoints = padding.LeftPoints;
                    fillTarget.PaddingTopPoints = padding.TopPoints;
                    fillTarget.PaddingRightPoints = padding.RightPoints;
                    fillTarget.PaddingBottomPoints = padding.BottomPoints;
                }
                if (source.GetVerticalAlignment(row, column) is { } vertical) fillTarget.VerticalAlignment = vertical switch {
                    IWorkCellVerticalAlignment.Top => PowerPointTextVerticalAlignment.Top,
                    IWorkCellVerticalAlignment.Middle => PowerPointTextVerticalAlignment.Center,
                    _ => PowerPointTextVerticalAlignment.Bottom
                };
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
            IWorkParagraphStyle? defaultStyle = source.GetParagraphStyle(sourceCell.Row, sourceCell.Column);
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
                    WriteParagraphContent(paragraphs[index], sourceParagraph, cancellationToken, resolveHyperlink, defaultStyle?.TextStyle);
                }
            } else {
                sourceCell.TryGetFormattedNumber(out string cellText, out string? numericColor);
                target.Text = cellText;
                if (defaultStyle != null) {
                    foreach (PowerPointParagraph paragraph in target.Paragraphs) {
                        ApplyParagraphStyle(paragraph, defaultStyle, paragraph.Text);
                        foreach (PowerPointTextRun run in paragraph.Runs) ApplyTextStyle(run, defaultStyle.TextStyle);
                    }
                }
                if (numericColor != null) {
                    foreach (PowerPointTextRun run in target.Paragraphs.SelectMany(paragraph => paragraph.Runs)) run.Color = numericColor;
                }
            }
        }
        foreach (IWorkTableMergeRange merge in source.MergedRanges) {
            cancellationToken.ThrowIfCancellationRequested();
            table.MergeCells(merge.FirstRow - 1, merge.FirstColumn - 1,
                merge.LastRow - 1, merge.LastColumn - 1);
        }
        ApplyAxisSizing(source.ColumnCount, source.ColumnWidths, source.DefaultColumnWidth,
            source.Geometry?.WidthPoints, width, (index, value) => table.SetColumnWidth(index, OfficeIMO.OfficeLength.Points(value)), cancellationToken);
        ApplyAxisSizing(source.RowCount, source.RowHeights, source.DefaultRowHeight,
            source.Geometry?.HeightPoints, height, (index, value) => table.SetRowHeight(index, OfficeIMO.OfficeLength.Points(value)), cancellationToken);
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
                || !FitsSpacing(style.SpaceBeforePoints) || !FitsSpacing(style.SpaceAfterPoints) || !FitsLineSpacing(style.LineSpacingMultiplier)
                || style.TabStops != null && style.TabStops.Any(tab => !FitsTextCoordinate(tab.PositionPoints)))
                return $"Keynote table '{table.Name}' contains paragraph formatting outside the PPTX range.";
            IWorkTextStyle text = style.TextStyle;
            if (!CanApplyCharacterTracking(text))
                return $"Keynote table '{table.Name}' contains character tracking without a qualified font size or outside the PPTX spacing range.";
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
