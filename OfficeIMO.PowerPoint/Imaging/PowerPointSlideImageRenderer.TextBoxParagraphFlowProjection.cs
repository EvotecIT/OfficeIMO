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
    private static bool TryAddTextBoxParagraphFlow(
        OfficeDrawing drawing,
        PowerPointTextBox textBox,
        double left,
        double top,
        double width,
        double height,
        double textWidth,
        double textHeight,
        double marginLeft,
        double marginTop,
        double rotation,
        double rotationCenterX,
        double rotationCenterY,
        bool flipHorizontal,
        bool flipVertical,
        PowerPointShapeBoundsMapping mapping,
        A.ColorScheme? colorScheme,
        List<OfficeImageExportDiagnostic> diagnostics, Func<string?, double, string?, OfficeFontStyle, double> measure, string? defaultFontFamily) {
        List<PowerPointParagraph> paragraphs = textBox.Paragraphs.ToList();
        if (!ShouldRenderTextBoxParagraphFlow(paragraphs)) {
            return false;
        }

        var numberingState = new PowerPointParagraphNumberingState();
        List<PowerPointParagraphDrawing> paragraphDrawings = CreateTextBoxParagraphDrawings(textBox, paragraphs, numberingState, textWidth, mapping, colorScheme, measure, defaultFontFamily);
        double flowHeight = paragraphDrawings.Sum(paragraph => paragraph.TotalHeight);
        double currentY = top + marginTop + ResolveTextBoxVerticalOffset(textBox.TextVerticalAlignment, textHeight, flowHeight);
        double contentBottom = top + marginTop + textHeight;

        for (int i = 0; i < paragraphDrawings.Count; i++) {
            PowerPointParagraphDrawing paragraph = paragraphDrawings[i];
            currentY += paragraph.SpaceBefore;
            if (!paragraph.HasContent) {
                // Empty paragraphs affect following origins but have no content to clip.
                currentY += paragraph.Height + paragraph.SpaceAfter;
                continue;
            }
            // Paragraph leading controls the next origin, not a separate clipping frame.
            // Condensed lines may paint below their final line box inside the text box.
            double visibleHeight = Math.Max(0D, contentBottom - currentY);
            bool clipped = visibleHeight + 0.001D < Math.Max(paragraph.Height, paragraph.PaintedHeight);
            if (clipped) {
                diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                    global::OfficeIMO.OfficeDiagnosticCodes.PowerpointTextOverflow, "Clipped PowerPoint paragraph content at the text frame boundary.", DescribeShape(textBox), OfficeConversionLossKind.Omission));
            }
            if (visibleHeight <= 0D) return true;

            if (paragraph.RichRuns.Count > 0) {
                drawing.AddRichText(
                    paragraph.RichRuns,
                    left + marginLeft,
                    currentY,
                    textWidth,
                    visibleHeight,
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
                    left + marginLeft,
                    currentY,
                    textWidth,
                    visibleHeight,
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

            if (visibleHeight < paragraph.Height) return true;
            currentY += paragraph.Height + paragraph.SpaceAfter;
        }

        return true;
    }

    private static bool ShouldRenderTextBoxParagraphFlow(IReadOnlyList<PowerPointParagraph> paragraphs) =>
        paragraphs.Any(paragraph => !string.IsNullOrEmpty(paragraph.BulletCharacter) || paragraph.IsNumbered ||
            paragraph.SpaceBeforePoints.HasValue || paragraph.SpaceAfterPoints.HasValue ||
            paragraph.LineSpacingPoints.HasValue || paragraph.LineSpacingMultiplier.HasValue) ||
        (paragraphs.Count > 1 &&
            paragraphs.Any(paragraph =>
                !paragraph.InlineNodes.Any(node => !string.IsNullOrEmpty(node.Text)) ||
                paragraph.Alignment != paragraphs[0].Alignment));
}
