namespace OfficeIMO.OpenDocument;

/// <summary>Reads and writes native uniform opacity values for graphic and text styles.</summary>
internal static class OdfOpacity {
    internal static double? ReadText(XElement? properties) {
        string? extension = (string?)properties?.Attribute(OdfNamespaces.LoExt + "opacity");
        string? legacy = (string?)properties?.Attribute(OdfNamespaces.Draw + "opacity");
        if (extension == null && legacy == null) return null;
        double value = Parse(extension ?? legacy!);
        if (extension != null && legacy != null && Parse(legacy) != value)
            throw new InvalidDataException("Conflicting native text opacity declarations.");
        return value;
    }

    internal static void Validate(double? value) {
        if (value.HasValue && (double.IsNaN(value.Value) || double.IsInfinity(value.Value) || value.Value < 0D || value.Value > 1D))
            throw new ArgumentOutOfRangeException(nameof(value), "Opacity must be between zero and one.");
    }

    // ODF percentage tokens do not accept exponent notation. Expand the round-trip
    // numeric representation rather than rounding small valid fractions to zero.
    internal static string? Format(double? value) {
        Validate(value);
        if (!value.HasValue) return null;
        if (value.Value == 0) return "0%";
        string percent = (value.Value * 100D).ToString("R", CultureInfo.InvariantCulture);
        int exponentIndex = percent.IndexOf('E');
        if (exponentIndex >= 0) {
            string mantissa = percent.Substring(0, exponentIndex);
            int exponent = int.Parse(percent.Substring(exponentIndex + 1), CultureInfo.InvariantCulture);
            int point = mantissa.IndexOf('.');
            if (point < 0) point = mantissa.Length;
            point += exponent;
            string digits = mantissa.Replace(".", string.Empty);
            percent = point <= 0 ? "0." + new string('0', -point) + digits :
                point >= digits.Length ? digits + new string('0', point - digits.Length) : digits.Insert(point, ".");
        }
        return percent + "%";
    }

    internal static double Parse(string value, bool allowFraction = false) {
        string lexical = value.Trim();
        bool percent = lexical.EndsWith("%", StringComparison.Ordinal);
        if (percent) lexical = lexical.Substring(0, lexical.Length - 1);
        if (percent && lexical.Any(character => character != '.' && (character < '0' || character > '9')))
            throw new InvalidDataException("Invalid percentage opacity value.");
        if ((!percent && !allowFraction) || !double.TryParse(lexical, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
            throw new InvalidDataException("Invalid native opacity value.");
        if (percent) number /= 100D;
        if (double.IsNaN(number) || double.IsInfinity(number) || number < 0D || number > 1D)
            throw new InvalidDataException("Native opacity must be between zero and one.");
        return number;
    }
}
