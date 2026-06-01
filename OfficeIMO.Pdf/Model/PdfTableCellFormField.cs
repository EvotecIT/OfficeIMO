namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a simple AcroForm text or choice field rendered inside a PDF table cell.
/// </summary>
public sealed class PdfTableCellFormField {
    private readonly PdfFormFieldStyle _style;

    private PdfTableCellFormField(
        PdfTableCellFormFieldKind kind,
        string name,
        double width,
        double height,
        double fontSize,
        string? value,
        System.Collections.Generic.IEnumerable<string>? options,
        bool isComboBox,
        PdfFormFieldStyle? style) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.Positive(fontSize, nameof(fontSize));

        Kind = kind;
        Name = name;
        Width = width;
        Height = height;
        FontSize = fontSize;
        IsComboBox = isComboBox;
        _style = style?.Clone() ?? new PdfFormFieldStyle();

        if (kind == PdfTableCellFormFieldKind.Text) {
            Value = value ?? string.Empty;
            Options = System.Array.AsReadOnly(System.Array.Empty<string>());
            Values = System.Array.AsReadOnly(new[] { Value });
            return;
        }

        System.Collections.Generic.List<string> normalizedOptions = NormalizeOptions(options);
        string selectedValue = string.IsNullOrWhiteSpace(value) ? normalizedOptions[0] : value!;
        if (!normalizedOptions.Contains(selectedValue, System.StringComparer.Ordinal)) {
            throw new System.ArgumentException("PDF table cell choice field value must match the provided options.", nameof(value));
        }

        Value = selectedValue;
        Values = System.Array.AsReadOnly(new[] { selectedValue });
        Options = normalizedOptions.AsReadOnly();
    }

    /// <summary>Kind of AcroForm field emitted for this table cell item.</summary>
    public PdfTableCellFormFieldKind Kind { get; }

    /// <summary>Field name written to the AcroForm tree.</summary>
    public string Name { get; }

    /// <summary>Initial scalar value.</summary>
    public string Value { get; }

    /// <summary>Initial selected values. Scalar table-cell fields expose one value.</summary>
    public System.Collections.Generic.IReadOnlyList<string> Values { get; }

    /// <summary>Available choice options. Empty for text fields.</summary>
    public System.Collections.Generic.IReadOnlyList<string> Options { get; }

    /// <summary>Preferred visual width in points. Rendering clamps this to the available cell width.</summary>
    public double Width { get; }

    /// <summary>Visual height in points.</summary>
    public double Height { get; }

    /// <summary>Text font size in points.</summary>
    public double FontSize { get; }

    /// <summary>Whether a choice field is emitted as a combo box.</summary>
    public bool IsComboBox { get; }

    /// <summary>Visual style for the generated field appearance stream.</summary>
    public PdfFormFieldStyle Style => _style.Clone();

    /// <summary>Creates a table-cell text field.</summary>
    public static PdfTableCellFormField TextField(string name, string? value = null, double width = 120, double height = 18, double fontSize = 10, PdfFormFieldStyle? style = null) =>
        new PdfTableCellFormField(PdfTableCellFormFieldKind.Text, name, width, height, fontSize, value, null, isComboBox: false, style);

    /// <summary>Creates a table-cell scalar choice field.</summary>
    public static PdfTableCellFormField ChoiceField(string name, System.Collections.Generic.IEnumerable<string> options, string? value = null, double width = 120, double height = 18, double fontSize = 10, bool isComboBox = true, PdfFormFieldStyle? style = null) =>
        new PdfTableCellFormField(PdfTableCellFormFieldKind.Choice, name, width, height, fontSize, value, options, isComboBox, style);

    internal PdfTableCellFormField Clone() =>
        new PdfTableCellFormField(Kind, Name, Width, Height, FontSize, Value, Options, IsComboBox, _style);

    private static System.Collections.Generic.List<string> NormalizeOptions(System.Collections.Generic.IEnumerable<string>? options) {
        Guard.NotNull(options, nameof(options));

        var normalized = new System.Collections.Generic.List<string>();
        var seen = new System.Collections.Generic.HashSet<string>(System.StringComparer.Ordinal);
        foreach (string? option in options!) {
            Guard.NotNullOrWhiteSpace(option, nameof(options));
            if (!seen.Add(option!)) {
                throw new System.ArgumentException("PDF table cell choice field options must be unique.", nameof(options));
            }

            normalized.Add(option!);
        }

        if (normalized.Count == 0) {
            throw new System.ArgumentException("PDF table cell choice field requires at least one option.", nameof(options));
        }

        return normalized;
    }
}

/// <summary>Supported table-cell AcroForm field kinds.</summary>
public enum PdfTableCellFormFieldKind {
    /// <summary>Simple text field.</summary>
    Text,

    /// <summary>Scalar choice field.</summary>
    Choice
}
