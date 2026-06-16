namespace OfficeIMO.Rtf;

/// <summary>
/// Raw shape instruction control from an RTF <c>\shpinst</c> destination.
/// </summary>
public sealed class RtfShapeInstruction {
    /// <summary>Creates a shape instruction control.</summary>
    public RtfShapeInstruction(string name, int? parameter = null, bool hasParameter = false) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Shape instruction name cannot be empty.", nameof(name));
        Name = name;
        Parameter = parameter;
        HasParameter = hasParameter;
    }

    /// <summary>Control word name without the leading backslash.</summary>
    public string Name { get; }

    /// <summary>Optional numeric control parameter.</summary>
    public int? Parameter { get; set; }

    /// <summary>Whether the control explicitly supplied a parameter.</summary>
    public bool HasParameter { get; set; }
}
