using System.Globalization;
using OfficeIMO.Core.Internal;
using OfficeIMO.Html;

namespace OfficeIMO.Epub.Image;

internal static class EpubFixedLayoutViewport {
    internal static bool TryRead(HtmlConversionDocument document, out double width, out double height) {
        width = height = 0D;
        var viewport = document.Document.QuerySelector("meta[name='viewport']");
        string? content = viewport?.GetAttribute("content");
        if (content != null) return TryRead(content, out width, out height);
        var svg = document.Document.QuerySelector("body > svg");
        if (svg == null) return false;
        return TryDimension(svg.GetAttribute("width"), out width) && TryDimension(svg.GetAttribute("height"), out height);
    }

    internal static bool TryRead(string content, out double width, out double height) {
        width = height = 0D;
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (string part in content.Split(new[] { ',', ';' })) {
            string[] pair = part.Split('=');
            if (pair.Length != 2 || values.ContainsKey(pair[0].Trim())) return false;
            values.Add(pair[0].Trim(), pair[1].Trim());
        }
        return values.TryGetValue("width", out string? widthText) && TryDimension(widthText, out width)
            && values.TryGetValue("height", out string? heightText) && TryDimension(heightText, out height);
    }

    private static bool TryDimension(string? value, out double result) =>
        OfficeRoundTripText.TryParse(value, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out result)
        && !double.IsNaN(result) && !double.IsInfinity(result) && result > 0D;
}
