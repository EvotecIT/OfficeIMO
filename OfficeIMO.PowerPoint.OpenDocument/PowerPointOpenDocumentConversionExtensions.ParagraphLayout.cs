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
                target.LineSpacingPoints = points;
            } else {
                unsupportedMeasurements++;
            }
        }
        return unsupportedMeasurements;
    }

    private static bool ApplyPowerPointParagraphLayout(PowerPointParagraph source, OdpParagraph target) {
        bool approximatedAlignment = false;
        if (source.RightToLeft == true) target.WritingMode = "rl-tb";
        switch (source.Alignment) {
            case PowerPointTextAlignment.Left:
                target.Alignment = OdpParagraphAlignment.Left;
                break;
            case PowerPointTextAlignment.Center:
                target.Alignment = OdpParagraphAlignment.Center;
                break;
            case PowerPointTextAlignment.Right:
                target.Alignment = OdpParagraphAlignment.Right;
                break;
            case PowerPointTextAlignment.Justified:
                target.Alignment = OdpParagraphAlignment.Justify;
                break;
            case PowerPointTextAlignment.JustifiedLow:
            case PowerPointTextAlignment.Distributed:
            case PowerPointTextAlignment.ThaiDistributed:
                target.Alignment = OdpParagraphAlignment.Justify;
                approximatedAlignment = true;
                break;
        }
        if (source.LineSpacingMultiplier.HasValue) {
            target.LineHeight = OdfLength.Parse(
                (source.LineSpacingMultiplier.Value * 100D).ToString("0.###", CultureInfo.InvariantCulture) + "%");
        } else if (source.LineSpacingPoints.HasValue) {
            target.LineHeight = OdfLength.Points(source.LineSpacingPoints.Value);
        }
        return approximatedAlignment;
    }
}
