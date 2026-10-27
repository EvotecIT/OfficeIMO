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
    private static void AddTextBox(OfficeDrawing drawing, PowerPointTextBox textBox, List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping, A.ColorScheme? colorScheme, bool suppressFrame = false, Func<string?, double, string?, OfficeFontStyle, double>? measure = null, string? defaultFontFamily = null, CancellationToken cancellationToken = default) {
        string text = ResolvePowerPointDisplayText(textBox);
        bool hasVisibleFrame = HasVisibleFrame(textBox, colorScheme);
        if (string.IsNullOrEmpty(text) && !hasVisibleFrame) {
            return;
        }

        if (!TryGetBounds(textBox, drawing, diagnostics, mapping, out double left, out double top, out double width, out double height)) {
            return;
        }

        OfficeShape frame = CreateTextBoxFrame(textBox, width, height, diagnostics);
        ApplyShapeStyle(frame, textBox, colorScheme, mapping, diagnostics);
        ApplyShapeTransform(frame, textBox, width, height);
        if (hasVisibleFrame && !suppressFrame) {
            drawing.AddShape(frame, left, top);
        }

        if (string.IsNullOrEmpty(text)) {
            return;
        }

        measure ??= OfficeDrawingTextLayout.CreateMetrics(drawing).MeasureText;
        OfficeDrawing textDrawing = CreateFittedTextBoxDrawing(textBox, width, height, diagnostics,
            mapping, colorScheme, measure, defaultFontFamily, cancellationToken);
        drawing.AddDrawingForClippedRendering(textDrawing, left, top, null);
    }

    private static OfficeShape CreateTextBoxFrame(
        PowerPointTextBox textBox,
        double width,
        double height,
        List<OfficeImageExportDiagnostic> diagnostics) {
        string? presetName = GetAutoShapePresetName(textBox);
        if (string.IsNullOrEmpty(presetName)
            || string.Equals(presetName, "rect", StringComparison.OrdinalIgnoreCase)
            || string.Equals(presetName, "rectangle", StringComparison.OrdinalIgnoreCase)) {
            return OfficeShape.Rectangle(width, height);
        }
        if (OfficeShapePresets.TryCreate(presetName, width, height,
                out OfficeShape? preset) && preset != null) {
            return preset;
        }

        AddUnsupportedShapeDiagnostic(
            diagnostics,
            textBox,
            "Rendered the PowerPoint text content inside a rectangular frame because its preset frame geometry is not yet projected through OfficeIMO.Drawing.",
            OfficeConversionLossKind.Approximation);
        return OfficeShape.Rectangle(width, height);
    }

    private static bool HasVisibleFrame(PowerPointShape source, A.ColorScheme? colorScheme) {
        return (source.FillTransparency != 100
                && (HasShapeFillGradient(source)
                    || TryResolveShapeFillColor(source, colorScheme, out _))) ||
            TryResolveShapeOutlineColor(source, colorScheme, out _);
    }

    private static OfficeFontInfo CreateFont(PowerPointTextBox textBox, PowerPointShapeBoundsMapping mapping, string? defaultFontFamily) {
        PowerPointTextRun? firstRun = textBox.Paragraphs
            .SelectMany(paragraph => paragraph.InlineNodes)
            .FirstOrDefault(node => node.Run != null && !string.IsNullOrEmpty(node.Text))?.Run;
        OfficeFontStyle style = OfficeFontStyle.Regular;
        if (firstRun?.Bold == true || textBox.Bold) {
            style |= OfficeFontStyle.Bold;
        }

        if (firstRun?.Italic == true || textBox.Italic) {
            style |= OfficeFontStyle.Italic;
        }

        return new OfficeFontInfo(
            firstRun?.FontName ?? textBox.FontName ?? ResolveTextBoxFallbackFont(textBox, defaultFontFamily),
            mapping.MapFontSize(firstRun?.FontSize ?? textBox.FontSize ?? 18),
            style);
    }

    private static List<OfficeRichTextRun> CreateRichTextRuns(PowerPointTextBox textBox, A.ColorScheme? colorScheme, PowerPointShapeBoundsMapping mapping, string? defaultFontFamily) {
        IReadOnlyList<PowerPointParagraph> paragraphs = textBox.Paragraphs;
        var richRuns = new List<OfficeRichTextRun>();
        var numberingState = new PowerPointParagraphNumberingState();
        for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++) {
            PowerPointParagraph paragraph = paragraphs[paragraphIndex];
            IReadOnlyList<PowerPointParagraphInline> inlineNodes = paragraph.InlineNodes;
            PowerPointTextRun? firstRun = inlineNodes.FirstOrDefault(node => node.Run != null)?.Run;
            string? marker = numberingState.CreateMarker(paragraph);
            bool paragraphHasVisibleText = !string.IsNullOrEmpty(marker) || inlineNodes.Any(node => !string.IsNullOrEmpty(node.Text));
            if (!paragraphHasVisibleText) {
                continue;
            }

            if (richRuns.Count > 0) {
                richRuns.Add(CreateRichTextRun(Environment.NewLine, firstRun, textBox, paragraph, colorScheme, mapping, defaultFontFamily: defaultFontFamily));
            }

            if (!string.IsNullOrEmpty(marker)) {
                richRuns.Add(CreateRichTextRun(marker!, firstRun, textBox, paragraph, colorScheme, mapping, markerRun: true, defaultFontFamily: defaultFontFamily));
            }

            for (int inlineIndex = 0; inlineIndex < inlineNodes.Count; inlineIndex++) {
                PowerPointParagraphInline inline = inlineNodes[inlineIndex];
                string runText = inline.Text;
                if (string.IsNullOrEmpty(runText)) {
                    continue;
                }

                richRuns.AddRange(CreateEffectiveTextRuns(runText, inline.Run, textBox, paragraph, colorScheme, mapping, defaultFontFamily));
            }
        }

        return richRuns;
    }

    private static OfficeTextVerticalAlignment MapTextVerticalAlignment(PowerPointTextVerticalAlignment? alignment) {
        if (alignment == PowerPointTextVerticalAlignment.Center) {
            return OfficeTextVerticalAlignment.Center;
        }

        if (alignment == PowerPointTextVerticalAlignment.Bottom) {
            return OfficeTextVerticalAlignment.Bottom;
        }

        return OfficeTextVerticalAlignment.Top;
    }
}
