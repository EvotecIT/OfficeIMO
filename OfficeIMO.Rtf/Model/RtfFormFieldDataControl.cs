namespace OfficeIMO.Rtf;

/// <summary>
/// Raw control word carried by an RTF <c>\ffdata</c> destination.
/// </summary>
public sealed class RtfFormFieldDataControl {
    /// <summary>Creates a form-field data control.</summary>
    public RtfFormFieldDataControl(string name, int? parameter = null, bool hasParameter = false) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Form-field data control name cannot be empty.", nameof(name));
        Name = name;
        Parameter = parameter;
        HasParameter = hasParameter;
    }

    /// <summary>Control word name without the leading backslash.</summary>
    public string Name { get; }

    /// <summary>Optional numeric parameter.</summary>
    public int? Parameter { get; set; }

    /// <summary>Whether the control explicitly supplied a parameter.</summary>
    public bool HasParameter { get; set; }
}
