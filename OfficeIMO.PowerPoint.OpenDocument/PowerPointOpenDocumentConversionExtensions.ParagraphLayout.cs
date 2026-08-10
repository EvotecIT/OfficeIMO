using System.Globalization;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.OpenDocument;

public static partial class PowerPointOpenDocumentConversionExtensions {
    private static int ApplyOdpParagraphLayout(
        OdpParagraph source,
        PowerPointParagraph target,
        ref int unsupportedWritingModes) {
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
                target.LineSpacingPoints = points;
            } else {
                unsupportedMeasurements++;
            }
        }
        return unsupportedMeasurements;
    }

    private static void ApplyPowerPointParagraphLayout(PowerPointParagraph source, OdpParagraph target) {
        if (source.RightToLeft == true) target.WritingMode = "rl-tb";
        if (source.LineSpacingMultiplier.HasValue) {
            target.LineHeight = OdfLength.Parse(
                (source.LineSpacingMultiplier.Value * 100D).ToString("0.###", CultureInfo.InvariantCulture) + "%");
        } else if (source.LineSpacingPoints.HasValue) {
            target.LineHeight = OdfLength.Points(source.LineSpacingPoints.Value);
        }
    }
}
