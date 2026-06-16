namespace OfficeIMO.Rtf;

/// <summary>
/// Custom document property from the RTF <c>\userprops</c> destination.
/// </summary>
public sealed class RtfUserProperty {
    /// <summary>Text property type code used by common RTF producers.</summary>
    public const int TextType = 30;

    /// <summary>Integer property type code used by common RTF producers.</summary>
    public const int IntegerType = 3;

    /// <summary>Floating-point property type code used by common RTF producers.</summary>
    public const int NumberType = 5;

    /// <summary>Boolean property type code used by common RTF producers.</summary>
    public const int BooleanType = 11;

    /// <summary>Date/time property type code used by common RTF producers.</summary>
    public const int DateTimeType = 64;

    /// <summary>Creates a custom property.</summary>
    public RtfUserProperty(string name, int? typeCode = null, string? staticValue = null) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Custom property name cannot be empty.", nameof(name));
        Name = name;
        TypeCode = typeCode;
        StaticValue = staticValue;
    }

    /// <summary>Property name from <c>\propname</c>.</summary>
    public string Name { get; set; }

    /// <summary>RTF property type code from <c>\proptype</c>.</summary>
    public int? TypeCode { get; set; }

    /// <summary>Static property value from <c>\staticval</c>.</summary>
    public string? StaticValue { get; set; }

    /// <summary>Linked property value from <c>\linkval</c>.</summary>
    public string? LinkedValue { get; set; }

    /// <summary>Creates a text custom property.</summary>
    public static RtfUserProperty Text(string name, string value) => new RtfUserProperty(name, TextType, value);

    /// <summary>Creates an integer custom property.</summary>
    public static RtfUserProperty Integer(string name, int value) => new RtfUserProperty(name, IntegerType, value.ToString(System.Globalization.CultureInfo.InvariantCulture));

    /// <summary>Creates a numeric custom property.</summary>
    public static RtfUserProperty Number(string name, double value) => new RtfUserProperty(name, NumberType, value.ToString(System.Globalization.CultureInfo.InvariantCulture));

    /// <summary>Creates a boolean custom property.</summary>
    public static RtfUserProperty Boolean(string name, bool value) => new RtfUserProperty(name, BooleanType, value ? "1" : "0");

    /// <summary>Creates a date/time custom property.</summary>
    public static RtfUserProperty DateTime(string name, System.DateTime value) => new RtfUserProperty(name, DateTimeType, value.ToString("O", System.Globalization.CultureInfo.InvariantCulture));
}
