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
        return WordVmlImageGeometry.ReadDimension(_vmlShape?.Style?.Value, name);
    }
}
