namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    private static Dictionary<string, PdfFormFieldValue> ToFormFieldValues(IReadOnlyDictionary<string, string> fieldValues) {
        Guard.NotNull(fieldValues, nameof(fieldValues));

        var converted = new Dictionary<string, PdfFormFieldValue>(fieldValues.Count, StringComparer.Ordinal);
        foreach (var pair in fieldValues) {
            if (pair.Value is null) {
                throw new ArgumentException("Field values cannot be null. Use an empty string to clear a text value.", nameof(fieldValues));
            }

            converted[pair.Key] = PdfFormFieldValue.From(pair.Value);
        }

        return converted;
    }

    private static void ValidateFieldValues(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        Guard.NotNull(fieldValues, nameof(fieldValues));
        if (fieldValues.Count == 0) {
            throw new ArgumentException("At least one field value is required.", nameof(fieldValues));
        }

        foreach (var pair in fieldValues) {
            if (string.IsNullOrWhiteSpace(pair.Key)) {
                throw new ArgumentException("Field names cannot be empty or whitespace.", nameof(fieldValues));
            }

            if (pair.Value is null) {
                throw new ArgumentException("Field values cannot be null. Use an empty string to clear a text value.", nameof(fieldValues));
            }

            if (pair.Value.Values.Count == 0) {
                throw new ArgumentException("Field values must contain at least one entry. Use an empty string to clear a text value.", nameof(fieldValues));
            }

            for (int i = 0; i < pair.Value.Values.Count; i++) {
                if (pair.Value.Values[i] is null) {
                    throw new ArgumentException("Field values cannot contain null entries. Use an empty string to clear a text value.", nameof(fieldValues));
                }
            }
        }
    }
}
