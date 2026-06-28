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
            A.ColorScheme? colorScheme,
            List<OfficeImageExportDiagnostic> diagnostics) {
            List<PowerPointParagraph> paragraphs = GetVisibleTableCellParagraphs(cell);
            if (!ShouldRenderTableCellParagraphFlow(paragraphs)) {
                return false;
            }

            var numberingState = new Dictionary<int, int>();
            List<PowerPointParagraphDrawing> paragraphDrawings = CreateTableCellParagraphDrawings(cell, paragraphs, numberingState, textWidth, colorScheme);
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
            (paragraphs.Count > 1 &&
                paragraphs.Any(paragraph =>
                    paragraph.SpaceBeforePoints.HasValue ||
                    paragraph.SpaceAfterPoints.HasValue ||
                    paragraph.LineSpacingPoints.HasValue ||
                    paragraph.LineSpacingMultiplier.HasValue));

        private static List<PowerPointParagraphDrawing> CreateTableCellParagraphDrawings(
            PowerPointTableCell cell,
            IReadOnlyList<PowerPointParagraph> paragraphs,
            Dictionary<int, int> numberingState,
            double textWidth,
            A.ColorScheme? colorScheme) {
            var results = new List<PowerPointParagraphDrawing>(paragraphs.Count);
            for (int i = 0; i < paragraphs.Count; i++) {
                PowerPointParagraph paragraph = paragraphs[i];
                string? marker = CreateParagraphMarker(paragraph, numberingState);
                OfficeTextParagraphIndent indent = CreateTableCellParagraphIndent(paragraph);
                OfficeTextAlignment alignment = MapTextAlignment(paragraph.Alignment ?? cell.HorizontalAlignment);
                List<OfficeRichTextRun> richRuns = CreateTableCellParagraphRichTextRuns(cell, paragraph, marker, colorScheme);
                double maxFontSize = richRuns.Count == 0
                    ? ResolveTableCellParagraphFont(cell, paragraph).Size
                    : richRuns.Max(run => run.FontSize);
                double lineHeight = ResolveTableCellParagraphLineHeight(paragraph, maxFontSize);
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
                        Math.Max(0D, paragraph.SpaceBeforePoints ?? 0D),
                        Math.Max(0D, paragraph.SpaceAfterPoints ?? 0D)));
                } else {
                    string text = CreateParagraphPlainText(paragraph, marker);
                    OfficeFontInfo font = ResolveTableCellParagraphFont(cell, paragraph);
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
                        Math.Max(0D, paragraph.SpaceBeforePoints ?? 0D),
                        Math.Max(0D, paragraph.SpaceAfterPoints ?? 0D)));
                }
            }

            return results;
        }

        private static List<OfficeRichTextRun> CreateTableCellParagraphRichTextRuns(PowerPointTableCell cell, PowerPointParagraph paragraph, string? marker, A.ColorScheme? colorScheme) {
            IReadOnlyList<PowerPointTextRun> runs = paragraph.Runs;
            var richRuns = new List<OfficeRichTextRun>(runs.Count);
            PowerPointTextRun? firstRun = runs.Count > 0 ? runs[0] : null;
            if (!string.IsNullOrEmpty(marker)) {
                richRuns.Add(CreateRichTextRun(marker!, firstRun, cell, paragraph, colorScheme, markerRun: true));
            }

            for (int i = 0; i < runs.Count; i++) {
                PowerPointTextRun run = runs[i];
                if (!string.IsNullOrEmpty(run.Text)) {
                    richRuns.Add(CreateRichTextRun(run.Text, run, cell, paragraph, colorScheme));
                }
            }

            return richRuns;
        }

        private static OfficeFontInfo ResolveTableCellParagraphFont(PowerPointTableCell cell, PowerPointParagraph paragraph) {
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

            return new OfficeFontInfo(firstRun?.FontName ?? cell.FontName ?? "Calibri", firstRun?.FontSize ?? cell.FontSize ?? 10, style);
        }

        private static OfficeColor ResolveTableCellParagraphTextColor(PowerPointTableCell cell, PowerPointParagraph paragraph, A.ColorScheme? colorScheme) {
            PowerPointTextRun? firstRun = paragraph.Runs.FirstOrDefault(run => !string.IsNullOrEmpty(run.Text)) ?? paragraph.Runs.FirstOrDefault();
            return ResolveTableCellTextRunColor(firstRun, cell, colorScheme);
        }

        private static double ResolveTableCellParagraphLineHeight(PowerPointParagraph paragraph, double fontSize) {
            if (paragraph.LineSpacingPoints.HasValue) {
                return Math.Max(1D, paragraph.LineSpacingPoints.Value);
            }

            if (paragraph.LineSpacingMultiplier.HasValue) {
                return Math.Max(1D, fontSize * paragraph.LineSpacingMultiplier.Value);
            }

            return Math.Max(1D, fontSize * 1.2D);
        }

        private static OfficeTextParagraphIndent CreateTableCellParagraphIndent(PowerPointParagraph paragraph) {
            double leftMargin = Math.Max(0D, paragraph.LeftMarginPoints ?? 0D);
            double firstLine = Math.Max(0D, leftMargin + (paragraph.IndentPoints ?? 0D));
            return firstLine > 0D || leftMargin > 0D
                ? new OfficeTextParagraphIndent(firstLine, leftMargin)
                : OfficeTextParagraphIndent.Empty;
        }
    }
}
