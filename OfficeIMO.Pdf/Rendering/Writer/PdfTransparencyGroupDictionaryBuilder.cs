using System.Globalization;
using System.Text;

namespace OfficeIMO.Pdf;

internal static class PdfTransparencyGroupDictionaryBuilder {
    internal static string BuildStreamDictionary(
        double width,
        double height,
        int contentLength,
        IReadOnlyList<(string Name, int Id)> fontResources,
        IReadOnlyList<(string Name, int Id)> xObjects,
        IReadOnlyList<(string Name, int Id)> graphicsStates,
        IReadOnlyList<(string Name, int Id)> shadings) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NonNegative(contentLength, nameof(contentLength));
        var resources = new StringBuilder();
        PdfPageDictionaryBuilder.AppendResourcePart(resources, "Font", fontResources);
        PdfPageDictionaryBuilder.AppendResourcePart(resources, "XObject", xObjects);
        PdfPageDictionaryBuilder.AppendResourcePart(resources, "ExtGState", graphicsStates);
        PdfPageDictionaryBuilder.AppendResourcePart(resources, "Shading", shadings);
        return "<< /Type /XObject /Subtype /Form /FormType 1 /BBox [0 0 "
            + Format(width) + " " + Format(height) + "] /Group << /S /Transparency /I true /K false >> /Resources <<"
            + resources + " >> /Length " + contentLength.ToString(CultureInfo.InvariantCulture) + " >>";
    }

    private static string Format(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);
}
