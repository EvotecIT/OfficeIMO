using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private bool TryDrawMixedText(
        string text,
        double x,
        double y,
        double width,
        double height,
        OfficeColor color,
        double fontSize,
        OfficeTextAlignment alignment,
        OfficeFontStyle style,
        string? fontFamily,
        OfficeTextOverflowBehavior overflowBehavior,
        double? textAdvanceWidth,
        OfficeTextDecorationStyle underlineStyle,
        OfficeTextDecorationStyle strikethroughStyle,
        OfficeColor? decorationColor,
        OfficeTextFeatureSettings? featureSettings,
        string? fontPalette,
        double? baselineFontSize,
        OfficeTextDirection textDirection) {
        if (_fonts == null) return false;
        IReadOnlyList<OfficeFontFallbackRun> runs = _fonts.PlanFallbackRuns(text, fontFamily, style);
        if (!ShouldUseFallbackRuns(runs, fontFamily)) return false;

        string value = text;
        bool retainOverflow = overflowBehavior == OfficeTextOverflowBehavior.Clip;
        double size = ResolveRasterTextSize(fontSize, height, textAdvanceWidth.HasValue);
        double availableWidth = Math.Max(1D, retainOverflow ? width : width - 6D);
        double measured = MeasurePositionedText(value, size, fontFamily, style, featureSettings, textDirection);
        if (!retainOverflow) {
            while (measured > availableWidth && value.Length > 0) {
                value = OfficeTextElements.RemoveLast(value);
                if (value.Length == 0) break;
                measured = MeasurePositionedText(value + "...", size, fontFamily, style, featureSettings, textDirection);
            }
            if (value.Length == 0 && MeasurePositionedText("...", size, fontFamily, style, featureSettings, textDirection) > availableWidth) return true;
            if (!string.Equals(value, text, StringComparison.Ordinal)) {
                value += "...";
                measured = MeasurePositionedText(value, size, fontFamily, style, featureSettings, textDirection);
                runs = _fonts.PlanFallbackRuns(value, fontFamily, style);
                if (!ShouldUseFallbackRuns(runs, fontFamily)) {
                    DrawTextCore(
                        value,
                        x,
                        y,
                        width,
                        height,
                        color,
                        fontSize,
                        alignment,
                        style,
                        fontFamily,
                        overflowBehavior,
                        textAdvanceWidth,
                        underlineStyle,
                        strikethroughStyle,
                        decorationColor,
                        featureSettings,
                        fontPalette,
                        baselineFontSize,
                        textDirection);
                    return true;
                }
            }
        }

        if (measured <= 0D) return true;
        double resolvedAdvance = textAdvanceWidth.HasValue && string.Equals(value, text, StringComparison.Ordinal)
            ? textAdvanceWidth.Value
            : measured;
        double textX = ResolveTextX(retainOverflow ? x : x + 3D, availableWidth, resolvedAdvance, alignment);
        double scale = resolvedAdvance / measured;
        double cursor = textX;
        foreach ((OfficeFontFallbackRun run, OfficeTextDirection runDirection) in PlanVisualFallbackRuns(value, fontFamily, style, textDirection)) {
            double runAdvance = MeasurePositionedText(run.Text, size, run.FamilyName, style, featureSettings, runDirection) * scale;
            DrawTextCore(
                run.Text,
                cursor,
                y,
                Math.Max(0.01D, runAdvance),
                height,
                color,
                size,
                OfficeTextAlignment.Left,
                style,
                run.FamilyName,
                OfficeTextOverflowBehavior.Clip,
                Math.Max(0.01D, runAdvance),
                underlineStyle,
                strikethroughStyle,
                decorationColor,
                featureSettings,
                fontPalette,
                baselineFontSize,
                runDirection);
            cursor += runAdvance;
        }
        return true;
    }

    private bool TryDrawMixedTextLine(
        string text,
        double anchorX,
        double top,
        double fontHeight,
        OfficeColor color,
        bool bold,
        bool italic,
        OfficeTextAlignment alignment,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        bool underline,
        bool strikethrough,
        string? fontFamily,
        bool flipHorizontal,
        bool flipVertical,
        OfficeTextDecorationStyle underlineStyle,
        OfficeTextDecorationStyle strikethroughStyle,
        OfficeColor? decorationColor) {
        if (_fonts == null) return false;
        OfficeFontStyle style = (bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
            | (italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
        IReadOnlyList<OfficeFontFallbackRun> runs = _fonts.PlanFallbackRuns(text, fontFamily, style);
        if (!ShouldUseFallbackRuns(runs, fontFamily)) return false;

        double width = MeasureText(text, fontHeight, fontFamily, style);
        double cursor = ResolveAnchoredTextX(anchorX, width, alignment);
        foreach (OfficeFontFallbackRun run in runs) {
            double runWidth = MeasureText(run.Text, fontHeight, run.FamilyName, style);
            DrawTextLine(
                run.Text,
                cursor,
                top,
                fontHeight,
                color,
                bold,
                italic,
                OfficeTextAlignment.Left,
                rotationDegrees,
                rotationCenterX,
                rotationCenterY,
                underline,
                strikethrough,
                run.FamilyName,
                flipHorizontal,
                flipVertical,
                underlineStyle,
                strikethroughStyle,
                decorationColor);
            cursor += runWidth;
        }
        return true;
    }

    private bool TryDrawMixedTransformedTextLine(
        string text,
        double anchorX,
        double top,
        double fontHeight,
        OfficeColor color,
        OfficeTransform transform,
        bool bold,
        bool italic,
        OfficeTextAlignment alignment,
        bool underline,
        bool strikethrough,
        string? fontFamily) {
        if (_fonts == null) return false;
        OfficeFontStyle style = (bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
            | (italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
        IReadOnlyList<OfficeFontFallbackRun> runs = _fonts.PlanFallbackRuns(text, fontFamily, style);
        if (!ShouldUseFallbackRuns(runs, fontFamily)) return false;

        double width = MeasureText(text, fontHeight, fontFamily, style);
        double cursor = ResolveAnchoredTextX(anchorX, width, alignment);
        foreach (OfficeFontFallbackRun run in runs) {
            double runWidth = MeasureText(run.Text, fontHeight, run.FamilyName, style);
            DrawTextLineTransformed(
                run.Text,
                cursor,
                top,
                fontHeight,
                color,
                transform,
                bold,
                italic,
                OfficeTextAlignment.Left,
                underline,
                strikethrough,
                run.FamilyName);
            cursor += runWidth;
        }
        return true;
    }

    private static bool ShouldUseFallbackRuns(
        IReadOnlyList<OfficeFontFallbackRun> runs,
        string? requestedFamilies) {
        if (runs.Count > 1) return true;
        if (runs.Count == 0) return false;
        return !string.Equals(
            runs[0].FamilyName,
            requestedFamilies?.Trim() ?? string.Empty,
            StringComparison.OrdinalIgnoreCase);
    }

    private IReadOnlyList<(OfficeFontFallbackRun Run, OfficeTextDirection Direction)> PlanVisualFallbackRuns(
        string text, string? fontFamily, OfficeFontStyle style, OfficeTextDirection textDirection) {
        var result = new List<(OfficeFontFallbackRun, OfficeTextDirection)>();
        // Resolve the complete string first. A font-only split does not know which
        // fallback face belongs at the visual left of an authored RTL run.
        foreach (OfficeBidiTextRun bidiRun in OfficeBidiTextResolver.ResolveVisualRuns(text, textDirection, _cancellationToken)) {
            IReadOnlyList<OfficeFontFallbackRun> faces = _fonts!.PlanFallbackRuns(bidiRun.Text, fontFamily, style);
            if (bidiRun.Direction == OfficeTextDirection.RightToLeft) {
                for (int index = faces.Count - 1; index >= 0; index--) result.Add((faces[index], bidiRun.Direction));
            } else {
                foreach (OfficeFontFallbackRun face in faces) result.Add((face, bidiRun.Direction));
            }
        }
        return result;
    }
}
