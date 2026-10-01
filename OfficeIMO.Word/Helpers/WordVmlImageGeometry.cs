using System.Globalization;

namespace OfficeIMO.Word;

/// <summary>Shared parsing of absolute VML image dimensions into the Word model's 96-DPI pixels.</summary>
internal static class WordVmlImageGeometry {
    internal static double? ReadDimension(string? style, string name) {
        foreach (string declaration in (style ?? string.Empty).Split(';')) {
            var pair = declaration.Split(':');
            if (pair.Length != 2 || !string.Equals(pair[0].Trim(), name, StringComparison.OrdinalIgnoreCase)) continue;
            string value = pair[1].Trim();
            double factor = 1;
            foreach (var unit in new[] { ("pt", 96d / 72d), ("px", 1d), ("in", 96d), ("cm", 96d / 2.54d), ("mm", 96d / 25.4d) }) {
                if (!value.EndsWith(unit.Item1, StringComparison.OrdinalIgnoreCase)) continue;
                factor = unit.Item2;
                value = value.Substring(0, value.Length - unit.Item1.Length);
                break;
            }
            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) &&
                number >= 0 && !double.IsInfinity(number) && !double.IsNaN(number)) return number * factor;
        }
        return null;
    }
}
