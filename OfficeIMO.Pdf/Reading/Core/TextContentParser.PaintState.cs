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
