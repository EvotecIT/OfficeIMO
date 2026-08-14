using System.Globalization;
using System.Text;

namespace OfficeIMO.Pdf;

internal static class PdfTransparencyGroupDictionaryBuilder {
    internal static string BuildStreamDictionary(
        double left,
        double bottom,
        double right,
        double top,
        int contentLength,
        IReadOnlyList<(string Name, int Id)> fontResources,
        IReadOnlyList<(string Name, int Id)> xObjects,
        IReadOnlyList<(string Name, int Id)> graphicsStates,
        IReadOnlyList<(string Name, int Id)> shadings,
        int? structParents = null) {
        ValidateFinite(left, nameof(left));
        ValidateFinite(bottom, nameof(bottom));
        ValidateFinite(right, nameof(right));
        ValidateFinite(top, nameof(top));
        if (right <= left) throw new ArgumentOutOfRangeException(nameof(right), "Transparency-group bounds must have positive width.");
        if (top <= bottom) throw new ArgumentOutOfRangeException(nameof(top), "Transparency-group bounds must have positive height.");
        Guard.NonNegative(contentLength, nameof(contentLength));
        var resources = new StringBuilder();
        PdfPageDictionaryBuilder.AppendResourcePart(resources, "Font", fontResources);
        PdfPageDictionaryBuilder.AppendResourcePart(resources, "XObject", xObjects);
        PdfPageDictionaryBuilder.AppendResourcePart(resources, "ExtGState", graphicsStates);
        PdfPageDictionaryBuilder.AppendResourcePart(resources, "Shading", shadings);
        string structureEntry = structParents.HasValue
            ? " /StructParents " + structParents.Value.ToString(CultureInfo.InvariantCulture)
            : string.Empty;
        return "<< /Type /XObject /Subtype /Form /FormType 1 /BBox ["
            + Format(left) + " " + Format(bottom) + " " + Format(right) + " " + Format(top)
            + "] /Group << /S /Transparency /I true /K false >> /Resources <<"
            + resources + " >>" + structureEntry + " /Length " + contentLength.ToString(CultureInfo.InvariantCulture) + " >>";
    }

    private static void ValidateFinite(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, value, "Transparency-group bounds must be finite.");
        }
    }

    private static string Format(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);
}
