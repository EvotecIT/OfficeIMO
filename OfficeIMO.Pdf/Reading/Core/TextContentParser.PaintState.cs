using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class TextContentParser {
    private static string BuildStrokeDashIdentity(List<object> operands) {
        if (operands.Count < 2) return "invalid";
        string values;
        if (operands[operands.Count - 2] is double[] numericValues) {
            values = string.Join(",", numericValues.Select(static number => number.ToString("R", CultureInfo.InvariantCulture)));
        } else if (operands[operands.Count - 2] is List<object> dashValues) {
            values = string.Join(",", dashValues.Select(static value => value is double number ? number.ToString("R", CultureInfo.InvariantCulture) : "invalid"));
        } else {
            return "invalid";
        }
        return values + ":" + (operands[operands.Count - 1] is double phase ? phase.ToString("R", CultureInfo.InvariantCulture) : "invalid");
    }

    private static bool IsSolidTextDash(string identity) => identity.Length > 0 && identity[0] == ':' ||
        identity.StartsWith("[]:", StringComparison.Ordinal);

    private static bool TryReadOutlineDash(string identity, double scale, out IReadOnlyList<double>? pattern, out double phase) {
        pattern = null; phase = 0D;
        if (IsSolidTextDash(identity)) return true;
        int separator = identity.IndexOf(':');
        if (separator <= 0) return false;
        if (!double.TryParse(
#if NET8_0_OR_GREATER
                identity.AsSpan(separator + 1),
#else
                identity.Substring(separator + 1),
#endif
                NumberStyles.Float,
                CultureInfo.InvariantCulture, out double rawPhase) || double.IsNaN(rawPhase) || double.IsInfinity(rawPhase)) return false;
        string[] values = identity.Substring(0, separator).Split(',');
        var dashes = new double[values.Length];
        bool positive = false;
        for (int index = 0; index < dashes.Length; index++) {
            if (!double.TryParse(values[index], NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
                value < 0D || double.IsNaN(value) || double.IsInfinity(value)) return false;
            dashes[index] = value * scale;
            if (double.IsInfinity(dashes[index])) return false;
            positive |= value > 0D;
        }
        phase = rawPhase * scale;
        if (!positive || double.IsInfinity(phase)) return false;
        pattern = Array.AsReadOnly(dashes);
        return true;
    }

    // A color-space selection resets its color, even without a following sc/SC.
    // Keep the authored initial components so later rendering intents can re-evaluate them.
    private static bool TryCreateInitialTextPaint(PdfPageColorSpace space, OfficeIccRenderingIntent intent,
        PdfOutputIntentColorTransform? outputTransform, out PdfPaintColorSelection? selection, out OfficeColor color) {
        selection = null;
        color = OfficeColor.Black;
        if (space.Kind == PdfPageColorSpaceKind.Pattern || space.SuppressesPaint) return false;
        var components = new object[space.ComponentCount];
        for (int index = 0; index < components.Length; index++) {
            components[index] = space.Kind is PdfPageColorSpaceKind.Separation or PdfPageColorSpaceKind.DeviceN ||
                space.Kind == PdfPageColorSpaceKind.DeviceCmyk && index == 3 ? 1D : 0D;
        }
        return PdfPaintColorSelection.TryCreate(components, space, intent, out selection, out color, outputTransform);
    }

}
