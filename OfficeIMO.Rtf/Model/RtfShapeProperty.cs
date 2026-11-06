namespace OfficeIMO.Rtf;

/// <summary>
/// Named shape property from an RTF <c>\sp</c> group.
/// </summary>
public sealed partial class RtfShapeProperty {
    /// <summary>Creates a named shape property.</summary>
    public RtfShapeProperty(string name, string? value = null) {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        Name = name;
        Value = value ?? string.Empty;
    }

    /// <summary>Property name from the <c>\sn</c> destination.</summary>
    public string Name { get; }

    /// <summary>Property value from the <c>\sv</c> destination.</summary>
    public string Value { get; set; }
}
