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
        double? textAdvanceWidth) {
        if (_fonts == null) return false;
        IReadOnlyList<OfficeFontFallbackRun> runs = _fonts.PlanFallbackRuns(text, fontFamily, style);
        if (runs.Count <= 1) return false;

        string value = text;
        bool retainOverflow = overflowBehavior == OfficeTextOverflowBehavior.Clip;
        double size = Math.Max(6D, Math.Min(fontSize, height - 2D));
        double availableWidth = Math.Max(1D, retainOverflow ? width : width - 6D);
        double measured = MeasureText(value, size, fontFamily, style);
        if (!retainOverflow) {
            while (measured > availableWidth && value.Length > 0) {
                value = OfficeTextElements.RemoveLast(value);
                if (value.Length == 0) break;
                measured = MeasureText(value + "...", size, fontFamily, style);
            }
            if (value.Length == 0 && MeasureText("...", size, fontFamily, style) > availableWidth) return true;
            if (!string.Equals(value, text, StringComparison.Ordinal)) {
                value += "...";
                measured = MeasureText(value, size, fontFamily, style);
                runs = _fonts.PlanFallbackRuns(value, fontFamily, style);
                if (runs.Count <= 1) {
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
                        textAdvanceWidth);
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
        foreach (OfficeFontFallbackRun run in runs) {
            double runAdvance = MeasureText(run.Text, size, run.FamilyName, style) * scale;
            DrawTextCore(
                run.Text,
                cursor,
                y,
                Math.Max(0.01D, runAdvance),
                height,
                color,
                fontSize,
                OfficeTextAlignment.Left,
                style,
                run.FamilyName,
                OfficeTextOverflowBehavior.Clip,
                Math.Max(0.01D, runAdvance));
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
        bool flipVertical) {
        if (_fonts == null) return false;
        OfficeFontStyle style = (bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
            | (italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
        IReadOnlyList<OfficeFontFallbackRun> runs = _fonts.PlanFallbackRuns(text, fontFamily, style);
        if (runs.Count <= 1) return false;

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
                flipVertical);
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
        if (runs.Count <= 1) return false;

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
}
