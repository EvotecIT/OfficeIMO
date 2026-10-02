using OfficeIMO.IWork;

namespace OfficeIMO.Word.IWork;

public static partial class WordIWorkConverter {
    private static string? FindWordProjectionLimitation(IWorkPagesProjection projection,
        bool allowPartialEditableReconstruction) {
        const long MaximumDestinationTableCells = 1_000_000;
        long destinationTableCells = 0;
        string? inlineLimitation = FindInlineObjectLimitation(projection);
        if (inlineLimitation != null) return inlineLimitation;
        if (projection.TextBoxObjects.Any(textBox => textBox.Hyperlink != null)
            || projection.Images.Any(image => image.Hyperlink != null)) {
            return "Pages contains a drawable hyperlink that cannot be represented by the DOCX owner.";
        }
        if (projection.Sections.SelectMany(section => section.HeaderContents)
            .Concat(projection.Sections.SelectMany(section => section.FooterContents))
            .Any(HasContainerScopedBreak)) {
            return "Pages headers or footers contain a section, layout, or page break that cannot be represented inside a DOCX header or footer.";
        }
        if (projection.TextBoxObjects.Any(textBox => HasContainerScopedBreak(textBox.Content))) {
            return "A Pages text box contains a section, layout, or page break that cannot be represented inside a DOCX text box.";
        }
        if (projection.Tables.SelectMany(table => table.Cells)
            .Any(cell => cell.RichText != null && HasContainerScopedBreak(cell.RichText))) {
            return "A Pages table cell contains a section, layout, or page break that cannot be represented inside a DOCX cell.";
        }
        if (projection.Body.Paragraphs
                .Concat(projection.TextBoxObjects.SelectMany(textBox => textBox.Content.Paragraphs))
                .Concat(projection.Sections.SelectMany(section => section.HeaderContents)
                    .Concat(projection.Sections.SelectMany(section => section.FooterContents))
                    .SelectMany(content => content.Paragraphs))
                .SelectMany(paragraph => paragraph.Runs)
                .Any(run => run.Hyperlink != null
                    && !Uri.TryCreate(run.Hyperlink, UriKind.Absolute, out _))) {
            return "Pages contains a text hyperlink that cannot be represented by the DOCX owner.";
        }
        if (projection.Tables.SelectMany(table => table.Cells)
            .Where(cell => cell.RichText != null)
            .SelectMany(cell => cell.RichText!.Paragraphs)
            .SelectMany(paragraph => paragraph.Runs)
            .Any(run => run.Hyperlink != null
                && !Uri.TryCreate(run.Hyperlink, UriKind.Absolute, out _))) {
            return "Pages contains a table-cell hyperlink that cannot be represented by the DOCX owner.";
        }
        if (AllPagesText(projection).SelectMany(content => content.Paragraphs)
                .Any(paragraph => paragraph.ListLevel > 8)) {
            return "Pages contains a list nesting level outside the DOCX numbering range.";
        }
        if (AllPagesText(projection).SelectMany(content => content.Paragraphs)
                .Any(paragraph => paragraph.ListLevel >= 0
                    && !IWorkNativeListCatalog.CanPreserveStart(paragraph.ListLabel))) {
            return "Pages contains an ordered-list marker that cannot be represented by DOCX numbering.";
        }
        if (projection.PageLayout is { } layout && !CanApplyPageLayout(layout, allowPartialEditableReconstruction)) {
            return "The Pages page layout exceeds the DOCX measurement range.";
        }
        var pagesWithBodyParagraphs = new HashSet<int>();
        int bodyPageIndex = 1;
        foreach (IWorkTextParagraph paragraph in projection.Body.Paragraphs) {
            pagesWithBodyParagraphs.Add(bodyPageIndex);
            if (paragraph.BreakKind is IWorkParagraphBreakKind.Page
                or IWorkParagraphBreakKind.Section) bodyPageIndex++;
        }
        if (projection.Drawables.Any(drawable => drawable.PageIndex.HasValue
                && !pagesWithBodyParagraphs.Contains(drawable.PageIndex.Value))) {
            return "A Pages drawable belongs to a source page with no DOCX anchor paragraph.";
        }
        foreach (IWorkTable table in projection.Tables) {
            if (table.Cells.Any(cell => cell.Comment is { } comment
                && (!WordCellComment.CanPreserveText(comment.Text) || !WordCellComment.CanPreserveText(comment.Author))))
                return $"Pages table '{table.Name}' contains comment text or authors that the DOCX owner cannot preserve without normalization.";
            if (!allowPartialEditableReconstruction && (table.HiddenRows.Count > 0 || table.HiddenColumns.Count > 0))
                return $"Pages table '{table.Name}' has hidden rows or columns that the DOCX table owner cannot preserve.";
            long tableCells = (long)table.RowCount * table.ColumnCount;
            if (table.RowCount == 0 || table.ColumnCount == 0) {
                return $"Pages table '{table.Name}' has no rows or columns and cannot be represented by the DOCX table owner.";
            }
            if (table.ColumnCount > 63) {
                return $"Pages table '{table.Name}' exceeds Word's supported 63-column table layout.";
            }
            if (table.RowCount > 32_767 || tableCells > 100_000) {
                return $"Pages table '{table.Name}' is too large for bounded DOCX table reconstruction.";
            }
            if (table.HasPopulatedCoveredMergeCells()) {
                return $"Pages table '{table.Name}' contains content in a covered merged cell that the DOCX owner cannot preserve.";
            }
            if (destinationTableCells > MaximumDestinationTableCells - tableCells) {
                return "Pages tables exceed the bounded DOCX destination cell budget.";
            }
            if (table.Cells.Any(cell => cell.Kind == IWorkCellKind.Formula
                && (cell.Value == null || !cell.CachedValueIsComplete))) {
                return $"Pages table '{table.Name}' contains a formula without a complete cached value that the DOCX owner cannot evaluate.";
            }
            if (table.Cells.Any(cell => cell.Kind == IWorkCellKind.Formula
                && cell.RichText is { IsComplete: false })) {
                return $"Pages table '{table.Name}' contains formula cached text with incomplete formatting that the DOCX owner cannot preserve.";
            }
            if (!FitsSignedTwips(table.DefaultRowHeight, allowPartialEditableReconstruction)
                || !FitsSignedTwips(table.DefaultColumnWidth, allowPartialEditableReconstruction)
                || table.RowHeights.Values.Any(height => !FitsSignedTwips(height, allowPartialEditableReconstruction))
                || table.ColumnWidths.Values.Any(width => !FitsSignedTwips(width, allowPartialEditableReconstruction))) {
                return $"Pages table '{table.Name}' has sizing outside the DOCX measurement range.";
            }
            if (table.Cells.Any(cell => cell.Padding != null
                && PaddingPoints(cell.Padding).Any(points => Math.Round(points * 20d, MidpointRounding.AwayFromZero) > short.MaxValue
                    || !allowPartialEditableReconstruction && !IsExactDestinationUnit(points, 20d)))) {
                return $"Pages table '{table.Name}' has cell padding outside the DOCX cell-margin range or twip precision.";
            }
            if (!allowPartialEditableReconstruction && table.Geometry is { } geometry
                && (Math.Abs(geometry.LeftPoints) > 0.000001d
                    || Math.Abs(geometry.TopPoints) > 0.000001d
                    || Math.Abs(geometry.WidthPoints) > 0.000001d
                    || Math.Abs(geometry.HeightPoints) > 0.000001d
                    || Math.Abs(geometry.RotationDegrees) > 0.000001d)) {
                return $"Pages table '{table.Name}' has positioned, sized, or rotated drawable geometry that the DOCX table owner cannot preserve.";
            }
            destinationTableCells += tableCells;
        }
        foreach (IWorkTextBox textBox in projection.TextBoxObjects) {
            if (textBox.Geometry is { } geometry
                && (!FitsEmuOffset(geometry.LeftPoints, allowPartialEditableReconstruction) || !FitsEmuOffset(geometry.TopPoints, allowPartialEditableReconstruction)
                    || !FitsEmuExtent(geometry.WidthPoints, allowPartialEditableReconstruction) || !FitsEmuExtent(geometry.HeightPoints, allowPartialEditableReconstruction)
                    || Math.Abs(geometry.RotationDegrees) > 0.000001d)) {
                return "A Pages text box has unsupported rotation or geometry outside the DOCX measurement range.";
            }
        }
        foreach (IWorkImageAsset image in projection.Images) {
            if (image.Geometry is { } geometry
                && (geometry.WidthPoints <= 0 || geometry.HeightPoints <= 0
                    || !FitsEmuExtent(geometry.WidthPoints, allowPartialEditableReconstruction)
                    || !FitsEmuExtent(geometry.HeightPoints, allowPartialEditableReconstruction)
                    || !FitsEmuOffset(geometry.LeftPoints, allowPartialEditableReconstruction)
                    || !FitsEmuOffset(geometry.TopPoints, allowPartialEditableReconstruction)
                    || Math.Abs(geometry.RotationDegrees) > 0.000001d)) {
                return "A Pages image has unsupported placement, rotation, or extent for the DOCX image owner.";
            }
        }
        foreach (IWorkParagraphStyle style in AllPagesParagraphStyles(projection)) {
            if (!FitsSignedTwips(style.FirstLineIndentPoints, allowPartialEditableReconstruction)
                || !FitsSignedTwips(style.LeftIndentPoints, allowPartialEditableReconstruction)
                || !FitsSignedTwips(style.RightIndentPoints, allowPartialEditableReconstruction)
                || !FitsUnsignedNullableTwips(style.SpaceBeforePoints, allowPartialEditableReconstruction)
                || !FitsUnsignedNullableTwips(style.SpaceAfterPoints, allowPartialEditableReconstruction)
                || style.TabStops != null && style.TabStops.Any(tab => !FitsSignedTwips(tab.PositionPoints, allowPartialEditableReconstruction))
                || style.LineSpacingMultiplier is double multiplier && (multiplier * 240d < 1d
                    || multiplier > int.MaxValue / 240d
                    || !allowPartialEditableReconstruction && !IsExactDestinationUnit(multiplier, 240d)))
                return "Pages paragraph formatting exceeds the DOCX measurement range.";
        }
        foreach (IWorkTextStyle style in AllPagesRunStyles(projection)) {
            if (style.Color is { Alpha: < byte.MaxValue } || style.BackgroundColor is { Alpha: < byte.MaxValue })
                return "Pages contains transparent text colors that cannot be represented by the DOCX owner.";
            if (style.FontSizePoints is double fontSize
                && (!IsFinite(fontSize) || fontSize < 0 || fontSize > int.MaxValue / 2d
                    || !allowPartialEditableReconstruction && fontSize * 2d != Math.Round(fontSize * 2d, MidpointRounding.AwayFromZero)))
                return "A Pages font size exceeds the DOCX measurement range or half-point precision.";
        }
        return null;
    }

    private static IEnumerable<double> PaddingPoints(IWorkCellPadding padding) {
        yield return padding.LeftPoints;
        yield return padding.TopPoints;
        yield return padding.RightPoints;
        yield return padding.BottomPoints;
    }

    private static bool RequiresWordRounding(IWorkPagesProjection projection) =>
        WordTwipMeasurements(projection).Any(value => value.HasValue && !IsExactDestinationUnit(value.Value, 20d))
        || AllPagesParagraphStyles(projection).Any(style => style.LineSpacingMultiplier is double multiplier
            && !IsExactDestinationUnit(multiplier, 240d))
        || AllPagesRunStyles(projection)
            .Any(style => style.FontSizePoints is double size && !IsExactDestinationUnit(size, 2d))
        || projection.TextBoxObjects.Select(box => box.Geometry)
            .Concat(projection.Images.Select(image => image.Geometry))
            .Any(geometry => geometry != null && (!IsExactDestinationUnit(geometry.LeftPoints, 12700d)
                || !IsExactDestinationUnit(geometry.TopPoints, 12700d)
                || !IsExactDestinationUnit(geometry.WidthPoints, 12700d)
                || !IsExactDestinationUnit(geometry.HeightPoints, 12700d)));

    private static IEnumerable<double?> WordTwipMeasurements(IWorkPagesProjection projection) {
        if (projection.PageLayout is { } layout) {
            yield return layout.WidthPoints;
            yield return layout.HeightPoints;
            yield return layout.LeftMarginPoints;
            yield return layout.RightMarginPoints;
            yield return layout.TopMarginPoints;
            yield return layout.BottomMarginPoints;
            yield return layout.HeaderMarginPoints;
            yield return layout.FooterMarginPoints;
        }
        foreach (IWorkTable table in projection.Tables) {
            yield return table.DefaultRowHeight;
            yield return table.DefaultColumnWidth;
            foreach (double height in table.RowHeights.Values) yield return height;
            foreach (double width in table.ColumnWidths.Values) yield return width;
            foreach (IWorkTableCell cell in table.Cells)
                if (cell.Padding != null)
                    foreach (double points in PaddingPoints(cell.Padding)) yield return points;
        }
        foreach (IWorkParagraphStyle style in AllPagesParagraphStyles(projection)) {
            yield return style.FirstLineIndentPoints;
            yield return style.LeftIndentPoints;
            yield return style.RightIndentPoints;
            yield return style.SpaceBeforePoints;
            yield return style.SpaceAfterPoints;
            if (style.TabStops != null) foreach (var tab in style.TabStops) yield return tab.PositionPoints;
        }
    }

    private static IEnumerable<IWorkParagraphStyle> TableParagraphStyles(IWorkTable table) =>
        new[] { table.TextStyles.Body, table.TextStyles.HeaderRow, table.TextStyles.HeaderColumn, table.TextStyles.FooterRow }
            .Concat(table.Cells.Select(cell => cell.ParagraphStyle)).Where(style => style != null).Cast<IWorkParagraphStyle>();

    private static IEnumerable<IWorkParagraphStyle> AllPagesParagraphStyles(IWorkPagesProjection projection) =>
        AllPagesText(projection).SelectMany(content => content.Paragraphs).Select(paragraph => paragraph.Style)
            .Concat(projection.Tables.SelectMany(TableParagraphStyles));

    private static IEnumerable<IWorkTextStyle> AllPagesRunStyles(IWorkPagesProjection projection) =>
        AllPagesText(projection).SelectMany(content => content.Paragraphs).SelectMany(paragraph => paragraph.Runs).Select(run => run.Style)
            .Concat(AllPagesText(projection).SelectMany(content => content.Paragraphs).Select(paragraph => paragraph.Style.TextStyle))
            .Concat(projection.Tables.SelectMany(TableParagraphStyles).Select(style => style.TextStyle));

    private static IEnumerable<IWorkTextContent> AllPagesText(IWorkPagesProjection projection) {
        yield return projection.Body;
        foreach (IWorkPagesSection section in projection.Sections) {
            foreach (IWorkTextContent header in section.HeaderContents) yield return header;
            foreach (IWorkTextContent footer in section.FooterContents) yield return footer;
        }
        foreach (IWorkTextBox textBox in projection.TextBoxObjects) yield return textBox.Content;
        foreach (IWorkTable table in projection.Tables) {
            foreach (IWorkTableCell cell in table.Cells) {
                if (cell.RichText != null) yield return cell.RichText;
            }
        }
    }

    private static bool HasContainerScopedBreak(IWorkTextContent content) =>
        content.Paragraphs.Any(paragraph => paragraph.BreakKind is IWorkParagraphBreakKind.Section
            or IWorkParagraphBreakKind.Layout or IWorkParagraphBreakKind.Page);

    private static bool FitsUnsignedTwips(double points, bool allowRounding = false) =>
        IsFinite(points) && points >= 0 && points <= uint.MaxValue / 20d
        && (allowRounding || IsExactDestinationUnit(points, 20d));

    private static bool CanApplyPageLayout(IWorkPageLayout layout, bool allowRounding = false) =>
        layout.WidthPoints > 0 && layout.HeightPoints > 0
        && layout.LeftMarginPoints + layout.RightMarginPoints < layout.WidthPoints
        && layout.TopMarginPoints + layout.BottomMarginPoints < layout.HeightPoints
        && FitsUnsignedTwips(layout.WidthPoints, allowRounding)
        && FitsUnsignedTwips(layout.HeightPoints, allowRounding)
        && FitsUnsignedTwips(layout.LeftMarginPoints, allowRounding)
        && FitsUnsignedTwips(layout.RightMarginPoints, allowRounding)
        && FitsSignedTwips(layout.TopMarginPoints, allowRounding)
        && FitsSignedTwips(layout.BottomMarginPoints, allowRounding)
        && FitsUnsignedTwips(layout.HeaderMarginPoints, allowRounding)
        && layout.HeaderMarginPoints <= layout.HeightPoints
        && FitsUnsignedTwips(layout.FooterMarginPoints, allowRounding)
        && layout.FooterMarginPoints <= layout.HeightPoints
        && (!allowRounding || RoundedPageLayoutFits(layout));

    private static bool RoundedPageLayoutFits(IWorkPageLayout layout) {
        double width = Math.Round(layout.WidthPoints * 20d, MidpointRounding.AwayFromZero);
        double height = Math.Round(layout.HeightPoints * 20d, MidpointRounding.AwayFromZero);
        double left = Math.Round(layout.LeftMarginPoints * 20d, MidpointRounding.AwayFromZero);
        double right = Math.Round(layout.RightMarginPoints * 20d, MidpointRounding.AwayFromZero);
        double top = Math.Round(layout.TopMarginPoints * 20d, MidpointRounding.AwayFromZero);
        double bottom = Math.Round(layout.BottomMarginPoints * 20d, MidpointRounding.AwayFromZero);
        return width > 0 && height > 0 && left + right < width && top + bottom < height;
    }

    private static bool FitsSignedTwips(double? points, bool allowRounding = false) => !points.HasValue
        || IsFinite(points.Value) && Math.Abs(points.Value) <= int.MaxValue / 20d
        && (allowRounding || IsExactDestinationUnit(points.Value, 20d));

    private static bool FitsUnsignedNullableTwips(double? points, bool allowRounding = false) => !points.HasValue
        || IsFinite(points.Value) && points.Value >= 0 && points.Value <= uint.MaxValue / 20d
        && (allowRounding || IsExactDestinationUnit(points.Value, 20d));

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static bool FitsEmuOffset(double points, bool allowRounding = false) =>
        !double.IsNaN(points) && !double.IsInfinity(points)
        && points >= int.MinValue / 12700d && points <= int.MaxValue / 12700d
        && (allowRounding || IsExactDestinationUnit(points, 12700d));

    private static bool FitsEmuExtent(double points, bool allowRounding = false) =>
        !double.IsNaN(points) && !double.IsInfinity(points)
        && points >= 0 && Math.Round(points * 12700d, MidpointRounding.AwayFromZero) < 9223372036854775808d
        && (allowRounding || IsExactDestinationUnit(points, 12700d));

    private static bool IsExactDestinationUnit(double points, double unitsPerPoint) {
        double scaled = points * unitsPerPoint;
        double rounded = Math.Round(scaled, MidpointRounding.AwayFromZero);
        double sourceFloatTolerance = Math.Max(1e-6d, Math.Abs(scaled) * 1e-7d);
        return Math.Abs(scaled - rounded) <= sourceFloatTolerance;
    }

    private static int ToEmusInt32(double points) => checked((int)Math.Round(points * 12700d,
        MidpointRounding.AwayFromZero));

    private static long ToEmusInt64(double points) => checked((long)Math.Round(points * 12700d,
        MidpointRounding.AwayFromZero));

}
