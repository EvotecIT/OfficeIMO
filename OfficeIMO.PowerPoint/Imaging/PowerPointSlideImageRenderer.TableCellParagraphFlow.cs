using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private static bool TryAddTableCellParagraphFlow(
            OfficeDrawing drawing,
            PowerPointTable table,
            PowerPointTableCell cell,
            double cellLeft,
            double cellTop,
            double cellWidth,
            double cellHeight,
            double textWidth,
            double textHeight,
            double marginLeft,
            double marginTop,
            double tableLeft,
            double tableTop,
            double tableWidth,
            double tableHeight,
            PowerPointShapeBoundsMapping mapping,
            A.ColorScheme? colorScheme,
            List<OfficeImageExportDiagnostic> diagnostics) {
            List<PowerPointParagraph> paragraphs = GetVisibleTableCellParagraphs(cell);
            if (!ShouldRenderTableCellParagraphFlow(paragraphs)) {
                return false;
            }

            var numberingState = new Dictionary<int, int>();
            List<PowerPointParagraphDrawing> paragraphDrawings = CreateTableCellParagraphDrawings(cell, paragraphs, numberingState, textWidth, mapping, colorScheme);
            double flowHeight = paragraphDrawings.Sum(paragraph => paragraph.TotalHeight);
            double currentY = cellTop + marginTop + ResolveTextBoxVerticalOffset(cell.VerticalAlignment, textHeight, flowHeight);
            double contentBottom = cellTop + marginTop + textHeight;
            double rotation = table.Rotation ?? 0D;
            double rotationCenterX = tableLeft + (tableWidth / 2D);
            double rotationCenterY = tableTop + (tableHeight / 2D);
            bool flipHorizontal = table.HorizontalFlip == true;
            bool flipVertical = table.VerticalFlip == true;

            for (int i = 0; i < paragraphDrawings.Count; i++) {
                PowerPointParagraphDrawing paragraph = paragraphDrawings[i];
                currentY += paragraph.SpaceBefore;
                if (currentY + paragraph.Height > contentBottom) {
                    AddUnsupportedShapeDiagnostic(diagnostics, table, "Skipped PowerPoint table cell paragraph content because it does not fit within the cell text frame.");
                    return true;
                }

                if (paragraph.RichRuns.Count > 0) {
                    drawing.AddRichText(
                        paragraph.RichRuns,
                        cellLeft + marginLeft,
                        currentY,
                        textWidth,
                        paragraph.Height,
                        paragraph.Alignment,
                        paragraph.LineHeight,
                        rotationDegrees: rotation,
                        rotationCenterX: rotationCenterX,
                        rotationCenterY: rotationCenterY,
                        wrapText: true,
                        flipHorizontal: flipHorizontal,
                        flipVertical: flipVertical,
                        paragraphIndent: paragraph.Indent);
                } else {
                    drawing.AddText(
                        paragraph.Text,
                        cellLeft + marginLeft,
                        currentY,
                        textWidth,
                        paragraph.Height,
                        paragraph.Font,
                        paragraph.Color,
                        paragraph.Alignment,
                        paragraph.LineHeight,
                        rotationDegrees: rotation,
                        rotationCenterX: rotationCenterX,
                        rotationCenterY: rotationCenterY,
                        wrapText: true,
                        flipHorizontal: flipHorizontal,
                        flipVertical: flipVertical,
                        paragraphIndent: paragraph.Indent);
                }

                currentY += paragraph.Height + paragraph.SpaceAfter;
            }

            return true;
        }

        private static List<PowerPointParagraph> GetVisibleTableCellParagraphs(PowerPointTableCell cell) =>
            cell.Cell.TextBody?
                .Elements<A.Paragraph>()
                .Select(paragraph => new PowerPointParagraph(paragraph))
                .Where(paragraph => paragraph.Runs.Any(run => !string.IsNullOrEmpty(run.Text)) || !string.IsNullOrEmpty(paragraph.BulletCharacter) || paragraph.IsNumbered)
                .ToList() ?? new List<PowerPointParagraph>();

        private static bool ShouldRenderTableCellParagraphFlow(IReadOnlyList<PowerPointParagraph> paragraphs) =>
            paragraphs.Any(paragraph => !string.IsNullOrEmpty(paragraph.BulletCharacter) || paragraph.IsNumbered) ||
            paragraphs.Count > 1;

        private static List<PowerPointParagraphDrawing> CreateTableCellParagraphDrawings(
            PowerPointTableCell cell,
            IReadOnlyList<PowerPointParagraph> paragraphs,
            Dictionary<int, int> numberingState,
            double textWidth,
            PowerPointShapeBoundsMapping mapping,
            A.ColorScheme? colorScheme) {
            var results = new List<PowerPointParagraphDrawing>(paragraphs.Count);
            for (int i = 0; i < paragraphs.Count; i++) {
                PowerPointParagraph paragraph = paragraphs[i];
                string? marker = CreateParagraphMarker(paragraph, numberingState);
                OfficeTextParagraphIndent indent = CreateTableCellParagraphIndent(paragraph, mapping);
                OfficeTextAlignment alignment = MapTextAlignment(paragraph.Alignment ?? cell.HorizontalAlignment);
                List<OfficeRichTextRun> richRuns = CreateTableCellParagraphRichTextRuns(cell, paragraph, marker, colorScheme, mapping);
                double maxFontSize = richRuns.Count == 0
                    ? ResolveTableCellParagraphFont(cell, paragraph, mapping).Size
                    : richRuns.Max(run => run.FontSize);
                double lineHeight = ResolveTableCellParagraphLineHeight(paragraph, maxFontSize, mapping);
                double height;
                if (ShouldRenderParagraphRichText(richRuns, marker)) {
                    height = EstimateParagraphRichTextHeight(richRuns, maxFontSize, lineHeight, textWidth, indent);
                    results.Add(new PowerPointParagraphDrawing(
                        string.Empty,
                        richRuns,
                        OfficeFontInfo.Default,
                        OfficeColor.Black,
                        alignment,
                        indent,
                        lineHeight,
                        height,
                        Math.Max(0D, mapping.MapVerticalLength(paragraph.SpaceBeforePoints ?? 0D)),
                        Math.Max(0D, mapping.MapVerticalLength(paragraph.SpaceAfterPoints ?? 0D))));
                } else {
                    string text = CreateParagraphPlainText(paragraph, marker);
                    OfficeFontInfo font = ResolveTableCellParagraphFont(cell, paragraph, mapping);
                    height = EstimateParagraphTextHeight(text, font, lineHeight, textWidth, indent);
                    results.Add(new PowerPointParagraphDrawing(
                        text,
                        Array.Empty<OfficeRichTextRun>(),
                        font,
                        ResolveTableCellParagraphTextColor(cell, paragraph, colorScheme),
                        alignment,
                        indent,
                        lineHeight,
                        height,
                        Math.Max(0D, mapping.MapVerticalLength(paragraph.SpaceBeforePoints ?? 0D)),
                        Math.Max(0D, mapping.MapVerticalLength(paragraph.SpaceAfterPoints ?? 0D))));
                }
            }

            return results;
        }

        private static List<OfficeRichTextRun> CreateTableCellParagraphRichTextRuns(PowerPointTableCell cell, PowerPointParagraph paragraph, string? marker, A.ColorScheme? colorScheme, PowerPointShapeBoundsMapping mapping) {
            IReadOnlyList<PowerPointTextRun> runs = paragraph.Runs;
            var richRuns = new List<OfficeRichTextRun>(runs.Count);
            PowerPointTextRun? firstRun = runs.Count > 0 ? runs[0] : null;
            if (!string.IsNullOrEmpty(marker)) {
                richRuns.Add(CreateRichTextRun(marker!, firstRun, cell, paragraph, colorScheme, mapping, markerRun: true));
            }

            for (int i = 0; i < runs.Count; i++) {
                PowerPointTextRun run = runs[i];
                if (!string.IsNullOrEmpty(run.Text)) {
                    richRuns.Add(CreateRichTextRun(run.Text, run, cell, paragraph, colorScheme, mapping));
                }
            }

            return richRuns;
        }

        private static OfficeFontInfo ResolveTableCellParagraphFont(PowerPointTableCell cell, PowerPointParagraph paragraph, PowerPointShapeBoundsMapping mapping) {
            PowerPointTextRun? firstRun = paragraph.Runs.FirstOrDefault(run => !string.IsNullOrEmpty(run.Text)) ?? paragraph.Runs.FirstOrDefault();
            OfficeFontStyle style = OfficeFontStyle.Regular;
            if (firstRun?.Bold == true || cell.Bold) {
                style |= OfficeFontStyle.Bold;
            }

            if (firstRun?.Italic == true || cell.Italic) {
                style |= OfficeFontStyle.Italic;
            }

            if (firstRun?.Underline == true) {
                style |= OfficeFontStyle.Underline;
            }

            if (firstRun?.Strikethrough == true) {
                style |= OfficeFontStyle.Strikethrough;
            }

            return new OfficeFontInfo(firstRun?.FontName ?? cell.FontName ?? "Calibri", mapping.MapFontSize(firstRun?.FontSize ?? cell.FontSize ?? 10), style);
        }

        private static OfficeColor ResolveTableCellParagraphTextColor(PowerPointTableCell cell, PowerPointParagraph paragraph, A.ColorScheme? colorScheme) {
            PowerPointTextRun? firstRun = paragraph.Runs.FirstOrDefault(run => !string.IsNullOrEmpty(run.Text)) ?? paragraph.Runs.FirstOrDefault();
            return ResolveTableCellTextRunColor(firstRun, cell, colorScheme);
        }

        private static double ResolveTableCellParagraphLineHeight(PowerPointParagraph paragraph, double fontSize, PowerPointShapeBoundsMapping mapping) {
            if (paragraph.LineSpacingPoints.HasValue) {
                return Math.Max(1D, mapping.MapVerticalLength(paragraph.LineSpacingPoints.Value));
            }

            if (paragraph.LineSpacingMultiplier.HasValue) {
                return Math.Max(1D, fontSize * paragraph.LineSpacingMultiplier.Value);
            }

            return Math.Max(1D, fontSize * 1.2D);
        }

        private static OfficeTextParagraphIndent CreateTableCellParagraphIndent(PowerPointParagraph paragraph, PowerPointShapeBoundsMapping mapping) {
            double leftMargin = Math.Max(0D, mapping.MapHorizontalLength(paragraph.LeftMarginPoints ?? 0D));
            double firstLine = Math.Max(0D, leftMargin + mapping.MapHorizontalLength(paragraph.IndentPoints ?? 0D));
            return firstLine > 0D || leftMargin > 0D
                ? new OfficeTextParagraphIndent(firstLine, leftMargin)
                : OfficeTextParagraphIndent.Empty;
        }
    }
}
