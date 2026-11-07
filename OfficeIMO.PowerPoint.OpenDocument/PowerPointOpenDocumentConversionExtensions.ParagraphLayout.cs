using System.Globalization;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.OpenDocument;

public static partial class PowerPointOpenDocumentConversionExtensions {
    private static int ApplyOdpParagraphLayout(
        OdpParagraph source,
        PowerPointParagraph target,
        ref int unsupportedWritingModes,
        ref int approximatedParagraphAlignments) {
        int unsupportedMeasurements = 0;
        string? writingMode = source.WritingMode;
        if (string.Equals(writingMode, "rl", StringComparison.OrdinalIgnoreCase)
            || string.Equals(writingMode, "rl-tb", StringComparison.OrdinalIgnoreCase)) {
            target.RightToLeft = true;
        } else if (string.Equals(writingMode, "lr", StringComparison.OrdinalIgnoreCase)
            || string.Equals(writingMode, "lr-tb", StringComparison.OrdinalIgnoreCase)) {
            target.RightToLeft = false;
        } else if (!string.IsNullOrWhiteSpace(writingMode)) {
            unsupportedWritingModes++;
        }

        switch (source.Alignment) {
            case OdpParagraphAlignment.Start:
                approximatedParagraphAlignments++;
                target.Alignment = source.IsRightToLeft
                    ? PowerPointTextAlignment.Right
                    : PowerPointTextAlignment.Left;
                break;
            case OdpParagraphAlignment.End:
                approximatedParagraphAlignments++;
                target.Alignment = source.IsRightToLeft
                    ? PowerPointTextAlignment.Left
                    : PowerPointTextAlignment.Right;
                break;
            case OdpParagraphAlignment.Left:
                target.Alignment = PowerPointTextAlignment.Left;
                break;
            case OdpParagraphAlignment.Center:
                target.Alignment = PowerPointTextAlignment.Center;
                break;
            case OdpParagraphAlignment.Right:
                target.Alignment = PowerPointTextAlignment.Right;
                break;
            case OdpParagraphAlignment.Justify:
                target.Alignment = PowerPointTextAlignment.Justified;
                break;
        }

        if (source.LineHeight.HasValue) {
            string lexical = source.LineHeight.Value.ToString();
            if (string.Equals(lexical, "normal", StringComparison.OrdinalIgnoreCase)) {
                // PowerPoint's absent line-spacing element carries the normal producer default.
            } else if (lexical.EndsWith("%", StringComparison.Ordinal)
                && double.TryParse(lexical.Substring(0, lexical.Length - 1), NumberStyles.Float,
                    CultureInfo.InvariantCulture, out double percentage)
                && percentage > 0D) {
                target.LineSpacingMultiplier = percentage / 100D;
            } else if (source.LineHeight.Value.TryToPoints(out double points) && points >= 0D) {
                target.LineSpacing = global::OfficeIMO.OfficeLength.Points(points);
            } else {
                unsupportedMeasurements++;
            }
        }
        return unsupportedMeasurements;
    }

    private static bool ApplyPowerPointParagraphLayout(PowerPointParagraph source, OdpParagraph target) {
        bool approximatedAlignment = false;
        OdfStyle? style = null;
        if (source.RightToLeft == true) {
            style = target.EnsureStyle();
            style.WritingMode = "rl-tb";
        }
        string? alignment = null;
        switch (source.Alignment) {
            case PowerPointTextAlignment.Left:
                alignment = "left";
                break;
            case PowerPointTextAlignment.Center:
                alignment = "center";
                break;
            case PowerPointTextAlignment.Right:
                alignment = "right";
                break;
            case PowerPointTextAlignment.Justified:
                alignment = "justify";
                break;
            case PowerPointTextAlignment.JustifiedLow:
            case PowerPointTextAlignment.Distributed:
            case PowerPointTextAlignment.ThaiDistributed:
                alignment = "justify";
                approximatedAlignment = true;
                break;
        }
        if (alignment != null) {
            style ??= target.EnsureStyle();
            style.TextAlign = alignment;
        }
        if (source.LineSpacingMultiplier.HasValue) {
            style ??= target.EnsureStyle();
            style.LineHeight = OdfLength.Parse(
                (source.LineSpacingMultiplier.Value * 100D).ToString("0.###", CultureInfo.InvariantCulture) + "%");
        } else if (source.LineSpacing.HasValue) {
            style ??= target.EnsureStyle();
            style.LineHeight = OdfLength.Points(source.LineSpacing.Value.ToPoints());
        }
        return approximatedAlignment;
    }
}
