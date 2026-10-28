using System.Globalization;

namespace OfficeIMO.Word;

/// <summary>Recognizes plain numeric values for the Word table authoring defaults.</summary>
internal static class WordNumericText {
    internal static bool IsNumber(string? text) {
        if (string.IsNullOrWhiteSpace(text)) return false;
        string value = text!.Trim().TrimEnd('%');
        return decimal.TryParse(value, NumberStyles.Number, CultureInfo.CurrentCulture, out _)
            || decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out _);
    }
}
