using System.Globalization;
using AngleSharp.Dom;
using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class HtmlExcelConverterExtensions {
    private static bool TryReadImagePosition(IElement item, string name, out double fallbackPixels) {
        bool hasFallback = TryReadDoubleAttribute(item, "data-officeimo-" + name, out fallbackPixels);
        return long.TryParse(item.GetAttribute("data-officeimo-" + name + "-emu"), NumberStyles.Integer,
            CultureInfo.InvariantCulture, out _) || hasFallback;
    }

    // Exact native metadata takes precedence over presentation pixels; both obey the import geometry budget.
    private static OfficeLength ReadImageLength(IElement item, string name, double fallbackPixels,
        bool positive, HtmlImportBudget budget, HtmlToExcelResult result, string kind = "image") {
        double maximum = Math.Min(int.MaxValue, budget.Limits.MaxAbsoluteGeometry);
        string? text = item.GetAttribute("data-officeimo-" + name + "-emu");
        if (text != null) {
            decimal limit = (decimal)maximum * 9525M;
            if (long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emus)
                && (!positive || emus > 0L) && emus <= limit && emus >= -limit)
                return OfficeLength.Emus(emus);
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticValueInvalid,
                "Invalid or out-of-range " + kind + " " + name + " native metadata used its pixel fallback.");
        }
        double fallback = positive ? (name == "height" ? 32D : 96D) : 0D;
        double pixels = NormalizeEditableLayoutGeometry(fallbackPixels, fallback,
            positive ? 1D / 9525D : -maximum, maximum, budget, result, kind + " " + name);
        return OfficeLength.Pixels(pixels);
    }
}
