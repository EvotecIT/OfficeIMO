namespace OfficeIMO.Pdf;

/// <summary>
/// Represents one or more values assigned to an AcroForm field.
/// </summary>
public sealed class PdfFormFieldValue {
    private readonly string[] _values;

    private PdfFormFieldValue(string[] values) {
        _values = values;
    }

    /// <summary>
    /// Values to store in the form field. Multiple values are used for simple multi-select choice fields.
    /// </summary>
    public IReadOnlyList<string> Values => _values;

    /// <summary>
    /// True when the field value contains more than one item.
    /// </summary>
    public bool IsMultiple => _values.Length > 1;

    /// <summary>
    /// Creates a scalar field value. Use an empty string to clear a simple text value.
    /// </summary>
    public static PdfFormFieldValue From(string value) {
        Guard.NotNull(value, nameof(value));
        return new PdfFormFieldValue(new[] { value });
    }

    /// <summary>
    /// Creates a field value containing one or more entries, primarily for simple multi-select choice fields.
    /// </summary>
    public static PdfFormFieldValue FromValues(params string[] values) {
        Guard.NotNull(values, nameof(values));
        return FromValues((IEnumerable<string>)values);
    }

    /// <summary>
    /// Creates a field value containing one or more entries, primarily for simple multi-select choice fields.
    /// </summary>
    public static PdfFormFieldValue FromValues(IEnumerable<string> values) {
        Guard.NotNull(values, nameof(values));

        string[] items = values.ToArray();
        if (items.Length == 0) {
            throw new ArgumentException("At least one field value is required. Use an empty string to clear a simple text value.", nameof(values));
        }

        for (int i = 0; i < items.Length; i++) {
            if (items[i] is null) {
                throw new ArgumentException("Field values cannot contain null entries. Use an empty string to clear a simple text value.", nameof(values));
            }
        }

        return new PdfFormFieldValue(items);
    }

    /// <summary>
    /// Converts a string into a scalar form field value.
    /// </summary>
    public static implicit operator PdfFormFieldValue(string value) => From(value);
}
