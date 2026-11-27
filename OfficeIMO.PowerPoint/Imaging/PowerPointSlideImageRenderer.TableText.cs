using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint;

internal static partial class PowerPointSlideImageRenderer {
    private static void AddTableCellText(
        OfficeDrawing drawing,
        PowerPointTable table,
        PowerPointTableCell cell,
        double cellLeft,
        double cellTop,
        double cellWidth,
        double cellHeight,
        double tableLeft,
        double tableTop,
        double tableWidth,
        double tableHeight,
        List<OfficeImageExportDiagnostic> diagnostics,
        PowerPointShapeBoundsMapping mapping,
        OfficeXmlNode colorScheme) {
        string text = ResolvePowerPointDisplayText(cell);
        if (string.IsNullOrEmpty(text)) {
            return;
        }

        AddSmallCapsApproximationDiagnosticIfNeeded(table, cell, diagnostics);
        AddWavyDoubleUnderlineApproximationDiagnosticIfNeeded(table, cell, diagnostics);
        AddBaselineApproximationDiagnosticIfNeeded(table, cell, diagnostics);

        double marginLeft = mapping.MapHorizontalLength(cell.PaddingLeft?.ToPoints() ?? 3.6D);
        double marginTop = mapping.MapVerticalLength(cell.PaddingTop?.ToPoints() ?? 1.8D);
        double marginRight = mapping.MapHorizontalLength(cell.PaddingRight?.ToPoints() ?? 3.6D);
        double marginBottom = mapping.MapVerticalLength(cell.PaddingBottom?.ToPoints() ?? 1.8D);
        double textWidth = cellWidth - marginLeft - marginRight;
        double textHeight = cellHeight - marginTop - marginBottom;
        if (textWidth <= 0D || textHeight <= 0D) {
            AddUnsupportedShapeDiagnostic(diagnostics, table, "Skipped PowerPoint table cell text because the cell margins leave no renderable drawing area.");
            return;
        }

        if (TryAddTableCellParagraphFlow(
            drawing,
            table,
            cell,
            cellLeft,
            cellTop,
            cellWidth,
            cellHeight,
            textWidth,
            textHeight,
            marginLeft,
            marginTop,
            tableLeft,
            tableTop,
            tableWidth,
            tableHeight,
            mapping,
            colorScheme,
            diagnostics)) {
            return;
        }

        var padding = new OfficeTextPadding(marginLeft, marginTop, marginRight, marginBottom);
        List<OfficeRichTextRun> richRuns = CreateRichTextRuns(cell, colorScheme, mapping);
        if (ShouldRenderRichText(richRuns)) {
            drawing.AddRichText(
                richRuns,
                cellLeft,
                cellTop,
                cellWidth,
                cellHeight,
                MapTextAlignment(cell.HorizontalAlignment),
                verticalAlignment: MapTextVerticalAlignment(cell.VerticalAlignment),
                rotationDegrees: table.Rotation ?? 0D,
                rotationCenterX: tableLeft + (tableWidth / 2D),
                rotationCenterY: tableTop + (tableHeight / 2D),
                wrapText: true,
                flipHorizontal: table.HorizontalFlip == true,
                flipVertical: table.VerticalFlip == true,
                padding: padding);
            return;
        }

        drawing.AddText(
            text,
            cellLeft,
            cellTop,
            cellWidth,
            cellHeight,
            CreateFont(cell, mapping),
            ResolveTableCellTextColor(cell, colorScheme),
            MapTextAlignment(cell.HorizontalAlignment),
            verticalAlignment: MapTextVerticalAlignment(cell.VerticalAlignment),
            rotationDegrees: table.Rotation ?? 0D,
            rotationCenterX: tableLeft + (tableWidth / 2D),
            rotationCenterY: tableTop + (tableHeight / 2D),
            wrapText: true,
            flipHorizontal: table.HorizontalFlip == true,
            flipVertical: table.VerticalFlip == true,
            padding: padding);
    }

    private static OfficeFontInfo CreateFont(PowerPointTableCell cell, PowerPointShapeBoundsMapping mapping) {
        PowerPointTextRun? firstRun = cell.Paragraphs
            .SelectMany(paragraph => paragraph.InlineNodes)
            .FirstOrDefault(node => node.Run != null && !string.IsNullOrEmpty(node.Text))?.Run;
        OfficeFontStyle style = OfficeFontStyle.Regular;
        if (firstRun?.Bold == true || cell.Bold) {
            style |= OfficeFontStyle.Bold;
        }

        if (firstRun?.Italic == true || cell.Italic) {
            style |= OfficeFontStyle.Italic;
        }

        return new OfficeFontInfo(
            ResolveThemeFontReference(firstRun?.FontFamily ?? cell.FontFamily,
                firstRun, null, cell) ?? "Calibri",
            mapping.MapFontSize(firstRun?.FontSize?.ToPoints() ?? cell.FontSize?.ToPoints() ?? 10),
            style);
    }

    private static List<OfficeRichTextRun> CreateRichTextRuns(PowerPointTableCell cell, OfficeXmlNode colorScheme, PowerPointShapeBoundsMapping mapping) {
        var richRuns = new List<OfficeRichTextRun>();
        if (cell.Element.Child(D, "txBody").IsNone) {
            return richRuns;
        }

        foreach (PowerPointParagraph paragraphModel in cell.Paragraphs) {
            IReadOnlyList<PowerPointParagraphInline> inlineNodes = paragraphModel.InlineNodes;
            bool paragraphHasVisibleText = inlineNodes.Any(node => !string.IsNullOrEmpty(node.Text));
            if (!paragraphHasVisibleText) {
                continue;
            }

            if (richRuns.Count > 0) {
                richRuns.Add(CreateRichTextRun(Environment.NewLine, null, cell, paragraphModel, colorScheme, mapping));
            }

            for (int inlineIndex = 0; inlineIndex < inlineNodes.Count; inlineIndex++) {
                PowerPointParagraphInline inline = inlineNodes[inlineIndex];
                string runText = inline.Text;
                if (string.IsNullOrEmpty(runText)) {
                    continue;
                }

                richRuns.Add(CreateRichTextRun(runText, inline.Run, cell, paragraphModel, colorScheme, mapping));
            }
        }

        return richRuns;
    }

    private static OfficeRichTextRun CreateRichTextRun(string text, PowerPointTextRun? run, PowerPointTableCell cell, OfficeXmlNode colorScheme, PowerPointShapeBoundsMapping mapping) {
        return CreateRichTextRun(text, run, cell, paragraph: null, colorScheme, mapping, markerRun: false);
    }

    private static OfficeRichTextRun CreateRichTextRun(string text, PowerPointTextRun? run, PowerPointTableCell cell, PowerPointParagraph? paragraph, OfficeXmlNode colorScheme, PowerPointShapeBoundsMapping mapping, bool markerRun = false) {
        OfficeColor color = ResolveTableCellTextRunColor(run, cell, colorScheme);
        if (markerRun && paragraph != null)
            color = ResolveTableCellBulletColor(paragraph, cell, colorScheme, color);
        OfficeColor? backgroundColor = ResolveTableCellTextRunBackgroundColor(run, cell, paragraph, colorScheme);
        Uri? hyperlink = markerRun ? null : run?.Hyperlink;
        return new OfficeRichTextRun(
            ResolvePowerPointDisplayText(text, run, paragraph, markerRun),
            mapping.MapFontSize(markerRun ? paragraph?.BulletSize?.ToPoints() ?? run?.FontSize?.ToPoints() ?? cell.FontSize?.ToPoints() ?? 10 : run?.FontSize?.ToPoints() ?? cell.FontSize?.ToPoints() ?? 10),
            color,
            run?.Bold == true,
            run?.Italic == true,
            run?.Underline == true,
            ResolveThemeFontReference(
                (markerRun ? paragraph?.BulletFontFamily : null) ?? run?.FontFamily ?? cell.FontFamily,
                run, paragraph, cell) ?? "Calibri",
            run?.Strikethrough == true,
            backgroundColor,
            MapUnderlineStyle(run?.UnderlineStyle),
            MapStrikeStyle(run?.StrikeStyle),
            MapBaseline(ResolvePowerPointBaselinePercent(run, paragraph))) {
            LinkUri = hyperlink is { IsAbsoluteUri: true } ? hyperlink.AbsoluteUri : null,
            LinkDestinationName = PowerPointHyperlinkResolver.TryParseSlideFragment(hyperlink, out int targetSlide)
                ? PowerPointHyperlinkResolver.GetSlideDestinationName(targetSlide) : null
        };
    }

    private static string ResolvePowerPointDisplayText(PowerPointTableCell cell) {
        IReadOnlyList<PowerPointParagraph> paragraphs = cell.Paragraphs;
        if (!paragraphs.Any(paragraph => paragraph.InlineNodes.Any(node => node.Run != null &&
                ResolvePowerPointCapitalization(node.Run, paragraph) is PowerPointCapitalization.AllCaps or PowerPointCapitalization.SmallCaps))) {
            return cell.Text;
        }

        return string.Join(Environment.NewLine, paragraphs.Select(paragraph =>
            string.Concat(paragraph.InlineNodes.Select(node =>
                node.Run == null ? node.Text : ResolvePowerPointDisplayText(node.Text, node.Run, paragraph)))));
    }

    private static void AddSmallCapsApproximationDiagnosticIfNeeded(
        PowerPointTable table,
        PowerPointTableCell cell,
        List<OfficeImageExportDiagnostic> diagnostics) {
        bool hasSmallCaps = cell.Paragraphs.Any(paragraph => paragraph.InlineNodes.Any(node => node.Run != null &&
            ResolvePowerPointCapitalization(node.Run, paragraph) == PowerPointCapitalization.SmallCaps && !string.IsNullOrEmpty(node.Text)));
        if (!hasSmallCaps) return;

        diagnostics.Add(new OfficeImageExportDiagnostic(
            OfficeImageExportDiagnosticSeverity.Warning,
            PowerPointImageExportDiagnosticCodes.SmallCapsApproximated,
            "Rendered native PowerPoint table-cell small caps as uppercase glyphs; reduced small-cap glyph sizing is approximated.",
            DescribeShape(table),
            OfficeConversionLossKind.Approximation));
    }

    private static void AddBaselineApproximationDiagnosticIfNeeded(
        PowerPointTable table,
        PowerPointTableCell cell,
        List<OfficeImageExportDiagnostic> diagnostics) {
        if (!cell.Paragraphs.Any(paragraph => paragraph.InlineNodes.Any(node => node.Run != null &&
                IsPowerPointBaselineApproximated(ResolvePowerPointBaselinePercent(node.Run, paragraph))))) {
            return;
        }

        diagnostics.Add(new OfficeImageExportDiagnostic(
            OfficeImageExportDiagnosticSeverity.Warning,
            PowerPointImageExportDiagnosticCodes.BaselinePercentApproximated,
            "Rendered a native PowerPoint table-cell baseline percentage using the shared superscript or subscript geometry; the exact authored percentage is not representable.",
            DescribeShape(table),
            OfficeConversionLossKind.Approximation));
    }

    private static void AddWavyDoubleUnderlineApproximationDiagnosticIfNeeded(
        PowerPointTable table,
        PowerPointTableCell cell,
        List<OfficeImageExportDiagnostic> diagnostics) {
        bool hasWavyDoubleUnderline = cell.Paragraphs
            .SelectMany(paragraph => paragraph.InlineNodes)
            .Any(node => node.Run?.UnderlineStyle == PowerPointUnderlineStyle.WavyDouble && !string.IsNullOrEmpty(node.Text));
        if (!hasWavyDoubleUnderline) return;

        diagnostics.Add(new OfficeImageExportDiagnostic(
            OfficeImageExportDiagnosticSeverity.Warning,
            PowerPointImageExportDiagnosticCodes.WavyDoubleUnderlineApproximated,
            "Rendered native PowerPoint table-cell double-wavy underline as a double solid underline because the shared Drawing model cannot represent the compound pattern.",
            DescribeShape(table),
            OfficeConversionLossKind.Approximation));
    }

    private static OfficeColor ResolveTableCellTextRunColor(PowerPointTextRun? run, PowerPointTableCell cell, OfficeXmlNode colorScheme) {
        OfficeColor? runColor = OfficeXmlThemeColors.ResolveColor(RunSolidFill(run), colorScheme,
            colorMap: GetSlideColorMap(TableCellSlide(cell)));
        if (runColor.HasValue) {
            return runColor.Value;
        }

        return ResolveTableCellTextColor(cell, colorScheme);
    }

    private static OfficeColor ResolveTableCellTextColor(PowerPointTableCell cell, OfficeXmlNode colorScheme) {
        PowerPointTextRun? run = cell.Paragraphs
            .SelectMany(paragraph => paragraph.InlineNodes)
            .FirstOrDefault(node => node.Run != null && !string.IsNullOrEmpty(node.Text))?.Run;
        OfficeColor? cellColor = OfficeXmlThemeColors.ResolveColor(RunSolidFill(run), colorScheme,
            colorMap: GetSlideColorMap(TableCellSlide(cell)));
        return cellColor.HasValue
            ? cellColor.Value
            : OfficeColor.Black;
    }

    /// <summary>The run's <c>a:rPr/a:solidFill</c>, or none.</summary>
    private static OfficeXmlNode RunSolidFill(PowerPointTextRun? run) =>
        run == null ? default : run.PropertiesElement.Child(D, "solidFill");

    /// <summary>The slide the cell's table is on, or none for a table of a layout or master.</summary>
    private static PowerPointPartRef TableCellSlide(PowerPointTableCell cell) {
        PowerPointPartRef slide = cell.SlidePartRef;
        return slide.Kind == PowerPointPartKind.Slide ? slide : default;
    }

    private static OfficeColor? ResolveTableCellTextRunBackgroundColor(PowerPointTextRun? run, PowerPointTableCell cell,
        PowerPointParagraph? paragraph, OfficeXmlNode colorScheme) {
        if (run != null) paragraph ??= FindRunParagraph(cell.Paragraphs, run);
        return OfficeXmlThemeColors.ResolveColor(
            ResolveTableCellHighlightElement(run, paragraph, cell), colorScheme,
            colorMap: GetSlideColorMap(TableCellSlide(cell)));
    }

    /// <summary>
    /// Resolves bullet paint of a table-cell paragraph from its own and its cell's list defaults, honoring an explicit
    /// return to text color.
    /// </summary>
    private static OfficeColor ResolveTableCellBulletColor(PowerPointParagraph paragraph, PowerPointTableCell cell,
        OfficeXmlNode colorScheme, OfficeColor textColor) {
        OfficeXmlNode paragraphProperties = paragraph.Element.Child(D, "pPr");
        int level = paragraphProperties.Int32Attribute("lvl") ?? 0;
        OfficeXmlNode listStyle = cell.Element.Child(D, "txBody").Child(D, "lstStyle");
        OfficeXmlNode colorMap = GetSlideColorMap(TableCellOwnerSlide(cell));
        foreach (OfficeXmlNode source in new[] {
                     paragraphProperties,
                     LevelParagraphProperties(listStyle, level),
                     listStyle.Child(D, "defPPr")
                 }) {
            if (!source.Child(D, "buClrTx").IsNone) return textColor;
            OfficeXmlNode bullet = source.Child(D, "buClr");
            if (!bullet.IsNone && OfficeXmlThemeColors.ResolveColor(bullet, colorScheme, colorMap: colorMap) is { } color) return color;
        }
        return textColor;
    }

    /// <summary>The part that holds the cell's table when it is a slide, or none.</summary>
    private static PowerPointPartRef TableCellOwnerSlide(PowerPointTableCell cell) {
        PowerPointPartRef owner = cell.OwnerPartRef;
        return owner.Kind == PowerPointPartKind.Slide ? owner : default;
    }

    /// <summary>
    /// The paragraph properties of a list style for a zero-based level (<c>a:lvl1pPr</c> to <c>a:lvl9pPr</c>), or none.
    /// </summary>
    private static OfficeXmlNode LevelParagraphProperties(OfficeXmlNode listStyle, int level) =>
        level is >= 0 and <= 8
            ? listStyle.Child(D, "lvl" + (level + 1).ToString(CultureInfo.InvariantCulture) + "pPr")
            : default;
}
