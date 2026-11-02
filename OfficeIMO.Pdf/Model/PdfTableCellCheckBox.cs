namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a simple AcroForm check box rendered inside a PDF table cell.
/// </summary>
public sealed class PdfTableCellCheckBox {
    private readonly PdfFormFieldStyle _style;

    /// <summary>Creates a table-cell check box.</summary>
    public PdfTableCellCheckBox(string name, bool isChecked = false, global::OfficeIMO.OfficeLength? size = null, string checkedValueName = "Yes", PdfFormFieldStyle? style = null) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Guard.Positive((size?.ToPoints() ?? 12), nameof(size));
        Guard.NotNullOrWhiteSpace(checkedValueName, nameof(checkedValueName));
        if (string.Equals(checkedValueName, "Off", System.StringComparison.Ordinal)) {
            throw new System.ArgumentException("Table cell check box selected value name cannot be Off.", nameof(checkedValueName));
        }

        Name = name;
        IsChecked = isChecked;
        Size = (size ?? global::OfficeIMO.OfficeLength.Points(12));
        CheckedValueName = checkedValueName;
        _style = style?.Clone() ?? new PdfFormFieldStyle();
    }

    /// <summary>Field name written to the AcroForm tree.</summary>
    public string Name { get; }

    /// <summary>Whether the generated check box is initially checked.</summary>
    public bool IsChecked { get; }

    /// <summary>Visual square size in points.</summary>
    public global::OfficeIMO.OfficeLength Size { get; }

    /// <summary>PDF button appearance state name used when checked.</summary>
    public string CheckedValueName { get; }

    /// <summary>Visual style for the generated check box appearance streams.</summary>
    public PdfFormFieldStyle Style => _style.Clone();

    internal PdfTableCellCheckBox Clone() => new PdfTableCellCheckBox(Name, IsChecked, Size, CheckedValueName, _style);
}
