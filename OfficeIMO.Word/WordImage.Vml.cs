using System.Globalization;

namespace OfficeIMO.Word;

public partial class WordImage {
    private void WriteVmlDimension(string name, double pixels) {
        if (double.IsNaN(pixels) || double.IsInfinity(pixels) || pixels < 0)
            throw new ArgumentOutOfRangeException(nameof(pixels));
        string replacement = name + ":" + pixels.ToString("R", CultureInfo.InvariantCulture) + "px";
        var declarations = new System.Collections.Generic.List<string>();
        bool replaced = false;
        foreach (string declaration in (_vmlShape?.Style?.Value ?? string.Empty).Split(';')) {
            if (string.IsNullOrWhiteSpace(declaration)) continue;
            int separator = declaration.IndexOf(':');
            if (separator >= 0 && string.Equals(declaration.Substring(0, separator).Trim(), name,
                    StringComparison.OrdinalIgnoreCase)) {
                if (!replaced) declarations.Add(replacement);
                replaced = true;
            } else declarations.Add(declaration.Trim());
        }
        if (!replaced) declarations.Add(replacement);
        _vmlShape!.Style = string.Join(";", declarations);
    }

    private double? ReadVmlDimension(string name) {
        foreach (string declaration in (_vmlShape?.Style?.Value ?? string.Empty).Split(';')) {
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
