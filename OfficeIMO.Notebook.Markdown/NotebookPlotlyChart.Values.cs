using System.Globalization;
using System.Text.Json;
using System.Xml;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private const int MaxLabelLength = 160;
    private const double MaxNumberMagnitude = 1e12;

    private static JsonElement Property(JsonElement source, string name) =>
        source.ValueKind == JsonValueKind.Object && source.TryGetProperty(name, out JsonElement value) ? value : default;

    private static bool HasProperty(JsonElement source, string name) => Property(source, name).ValueKind != JsonValueKind.Undefined;

    private static bool ContainsOnly(JsonElement source, params string[] allowed) {
        if (source.ValueKind == JsonValueKind.Undefined) return true;
        if (source.ValueKind != JsonValueKind.Object) return false;
        int count = 0;
        foreach (JsonProperty property in source.EnumerateObject()) {
            if (++count > allowed.Length) return false;
            bool known = false;
            foreach (string candidate in allowed) {
                if (property.NameEquals(candidate)) {
                    known = true;
                    break;
                }
            }
            if (!known) return false;
        }
        return true;
    }

    private static bool ValueOrDefault(JsonElement value, string expected, bool allowMissing = true) =>
        value.ValueKind == JsonValueKind.Undefined ? allowMissing :
        value.ValueKind == JsonValueKind.String && value.ValueEquals(expected);

    private static bool TryReadOptionalText(JsonElement source, out string? text) {
        text = null;
        if (source.ValueKind == JsonValueKind.Undefined) return true;
        if (!TryReadString(source, MaxLabelLength, out string? value) || value!.IndexOf('<') >= 0) return false;
        try {
            text = XmlConvert.VerifyXmlChars(value);
            return true;
        } catch (XmlException) {
            return false;
        }
    }

    private static bool TryReadString(JsonElement source, int limit, out string? value) {
        value = null;
        if (source.ValueKind != JsonValueKind.String) return false;
        try {
            value = source.GetString();
            return value != null && value.Length <= limit;
        } catch (InvalidOperationException) {
            return false;
        }
    }

    private static bool TryReadBoolean(JsonElement source, bool defaultValue, out bool value) {
        value = defaultValue;
        if (source.ValueKind == JsonValueKind.Undefined) return true;
        if (source.ValueKind != JsonValueKind.True && source.ValueKind != JsonValueKind.False) return false;
        value = source.GetBoolean();
        return true;
    }

    private static bool TryReadLength(JsonElement source, double minimum, double maximum, out OfficeLength? length) {
        length = null;
        if (source.ValueKind == JsonValueKind.Undefined) return true;
        if (!TryReadNumber(source, out double number) || number < minimum || number > maximum) return false;
        length = OfficeLength.Pixels(number);
        return true;
    }

    private static bool TryReadCategories(JsonElement source, int maxCount, bool explicitCategoryAxis,
        out string[]? categories) {
        categories = null;
        if (source.ValueKind != JsonValueKind.Array || source.GetArrayLength() == 0 || source.GetArrayLength() > maxCount)
            return false;
        var result = new string[source.GetArrayLength()];
        var distinct = new HashSet<string>(StringComparer.Ordinal);
        int index = 0;
        foreach (JsonElement category in source.EnumerateArray()) {
            if (!TryReadOptionalText(category, out string? label) || string.IsNullOrWhiteSpace(label) || !distinct.Add(label!))
                return false;
            // Plotly can infer a date or value axis from string arrays. Require an explicit categorical
            // axis for labels that we cannot safely treat as equally spaced without that declaration.
            if (!explicitCategoryAxis &&
                (double.TryParse(label, NumberStyles.Float, CultureInfo.InvariantCulture, out _) ||
                    DateTimeOffset.TryParse(label, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out _)))
                return false;
            result[index++] = label!;
        }
        categories = result;
        return true;
    }

    private static double[] DefaultScatterX(int count) {
        var result = new double[count];
        for (int index = 0; index < count; index++) result[index] = index;
        return result;
    }

    private static bool TryReadNumber(JsonElement element, out double value) {
        value = 0;
        return element.ValueKind == JsonValueKind.Number && element.TryGetDouble(out value) && IsSupportedNumber(value);
    }

    private static bool IsSupportedNumber(double value) => !double.IsNaN(value) && Math.Abs(value) <= MaxNumberMagnitude;

    private static double[]? ReadNumbers(JsonElement source, int maxCount) {
        if (maxCount <= 0) return null;
        if (source.ValueKind == JsonValueKind.Object) return ReadTypedNumbers(source, maxCount);
        if (source.ValueKind != JsonValueKind.Array || source.GetArrayLength() == 0 || source.GetArrayLength() > maxCount)
            return null;
        var values = new double[source.GetArrayLength()];
        int index = 0;
        foreach (JsonElement element in source.EnumerateArray()) {
            if (!TryReadNumber(element, out values[index++])) return null;
        }
        return values;
    }

    /// <summary>Reads the little-endian, one-dimensional NumPy array encoding used by Plotly.py.</summary>
    private static double[]? ReadTypedNumbers(JsonElement source, int maxCount) {
        if (!ContainsOnly(source, "dtype", "bdata") ||
            !TryReadString(Property(source, "dtype"), 2, out string? dtype)) return null;
        int stride = dtype switch {
            "i1" or "u1" => 1,
            "i2" or "u2" => 2,
            "i4" or "u4" or "f4" => 4,
            "f8" => 8,
            _ => 0
        };
        if (stride == 0 ||
            !TryReadString(Property(source, "bdata"), ((maxCount * stride + 2) / 3) * 4, out string? encoded))
            return null;
        byte[] bytes;
        try {
            bytes = Convert.FromBase64String(encoded!);
        } catch (FormatException) {
            return null;
        }
        int count = bytes.Length / stride;
        if (count == 0 || count > maxCount || bytes.Length % stride != 0) return null;
        var result = new double[count];
        for (int index = 0; index < count; index++) {
            int offset = index * stride;
            if (!BitConverter.IsLittleEndian && stride > 1) Array.Reverse(bytes, offset, stride);
            double value = dtype switch {
                "i1" => unchecked((sbyte)bytes[offset]),
                "u1" => bytes[offset],
                "i2" => BitConverter.ToInt16(bytes, offset),
                "u2" => BitConverter.ToUInt16(bytes, offset),
                "i4" => BitConverter.ToInt32(bytes, offset),
                "u4" => BitConverter.ToUInt32(bytes, offset),
                "f4" => BitConverter.ToSingle(bytes, offset),
                "f8" => BitConverter.ToDouble(bytes, offset),
                _ => double.NaN
            };
            if (!IsSupportedNumber(value)) return null;
            result[index] = value;
        }
        return result;
    }
}
