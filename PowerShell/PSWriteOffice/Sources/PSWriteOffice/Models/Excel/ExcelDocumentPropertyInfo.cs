namespace PSWriteOffice.Models.Excel;

/// <summary>Represents an Excel document property exposed to PowerShell.</summary>
public sealed class ExcelDocumentPropertyInfo
{
    /// <summary>Creates a new property descriptor.</summary>
    public ExcelDocumentPropertyInfo(string name, string scope, object? value, string? valueType, string? customPropertyType = null)
    {
        Name = name;
        Scope = scope;
        Value = value;
        ValueType = valueType;
        CustomPropertyType = customPropertyType;
    }

    /// <summary>Property name.</summary>
    public string Name { get; }

    /// <summary>Property scope (`BuiltIn`, `Application`, or `Custom`).</summary>
    public string Scope { get; }

    /// <summary>Property value.</summary>
    public object? Value { get; }

    /// <summary>.NET type name for <see cref="Value"/> when available.</summary>
    public string? ValueType { get; }

    /// <summary>Underlying OfficeIMO custom-property type, when applicable.</summary>
    public string? CustomPropertyType { get; }
}
