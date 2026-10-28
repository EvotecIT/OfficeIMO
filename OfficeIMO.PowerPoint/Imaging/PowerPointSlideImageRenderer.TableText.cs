using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

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
        A.ColorScheme? colorScheme) {
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
            firstRun?.FontName ?? cell.FontName ?? "Calibri",
            mapping.MapFontSize(firstRun?.FontSize?.ToPoints() ?? cell.FontSize?.ToPoints() ?? 10),
            style);
    }

    private static List<OfficeRichTextRun> CreateRichTextRuns(PowerPointTableCell cell, A.ColorScheme? colorScheme, PowerPointShapeBoundsMapping mapping) {
        var richRuns = new List<OfficeRichTextRun>();
        A.TextBody? textBody = cell.Cell.TextBody;
        if (textBody == null) {
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

    private static OfficeRichTextRun CreateRichTextRun(string text, PowerPointTextRun? run, PowerPointTableCell cell, A.ColorScheme? colorScheme, PowerPointShapeBoundsMapping mapping) {
        return CreateRichTextRun(text, run, cell, paragraph: null, colorScheme, mapping, markerRun: false);
    }

    private static OfficeRichTextRun CreateRichTextRun(string text, PowerPointTextRun? run, PowerPointTableCell cell, PowerPointParagraph? paragraph, A.ColorScheme? colorScheme, PowerPointShapeBoundsMapping mapping, bool markerRun = false) {
        OfficeColor color = ResolveTableCellTextRunColor(run, cell, colorScheme);
        OfficeColor? backgroundColor = ResolveTableCellTextRunBackgroundColor(run, colorScheme);
        Uri? hyperlink = markerRun ? null : run?.Hyperlink;
        return new OfficeRichTextRun(
            ResolvePowerPointDisplayText(text, run, paragraph, markerRun),
            mapping.MapFontSize(markerRun ? paragraph?.BulletSize?.ToPoints() ?? run?.FontSize?.ToPoints() ?? cell.FontSize?.ToPoints() ?? 10 : run?.FontSize?.ToPoints() ?? cell.FontSize?.ToPoints() ?? 10),
            color,
            run?.Bold == true,
            run?.Italic == true,
            run?.Underline == true,
            markerRun ? paragraph?.BulletFontName ?? run?.FontName ?? cell.FontName ?? "Calibri" : run?.FontName ?? cell.FontName ?? "Calibri",
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

    private static OfficeColor ResolveTableCellTextRunColor(PowerPointTextRun? run, PowerPointTableCell cell, A.ColorScheme? colorScheme) {
        OfficeColor? runColor = OfficeOpenXmlThemeColorResolver.ResolveColor(run?.RunProperties?.GetFirstChild<A.SolidFill>(), colorScheme);
        if (runColor.HasValue) {
            return runColor.Value;
        }

        return ResolveTableCellTextColor(cell, colorScheme);
    }

    private static OfficeColor ResolveTableCellTextColor(PowerPointTableCell cell, A.ColorScheme? colorScheme) {
        PowerPointTextRun? run = cell.Paragraphs
            .SelectMany(paragraph => paragraph.InlineNodes)
            .FirstOrDefault(node => node.Run != null && !string.IsNullOrEmpty(node.Text))?.Run;
        OfficeColor? cellColor = OfficeOpenXmlThemeColorResolver.ResolveColor(run?.RunProperties?.GetFirstChild<A.SolidFill>(), colorScheme);
        return cellColor.HasValue
            ? cellColor.Value
            : OfficeColor.Black;
    }

    private static OfficeColor? ResolveTableCellTextRunBackgroundColor(PowerPointTextRun? run, A.ColorScheme? colorScheme) {
        return OfficeOpenXmlThemeColorResolver.ResolveColor(run?.RunProperties?.GetFirstChild<A.Highlight>(), colorScheme);
    }
}
