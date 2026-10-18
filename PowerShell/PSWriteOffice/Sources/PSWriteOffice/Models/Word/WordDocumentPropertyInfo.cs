namespace PSWriteOffice.Models.Word;

/// <summary>Represents a Word document property exposed to PowerShell.</summary>
public sealed class WordDocumentPropertyInfo
{
    /// <summary>Creates a new property descriptor.</summary>
    public WordDocumentPropertyInfo(string name, string scope, object? value, string? valueType, string? customPropertyType)
    {
        Name = name;
        Scope = scope;
        Value = value;
        ValueType = valueType;
        CustomPropertyType = customPropertyType;
    }

    /// <summary>Property name.</summary>
    public string Name { get; }

    /// <summary>Property scope (`BuiltIn` or `Custom`).</summary>
    public string Scope { get; }

    /// <summary>Property value.</summary>
    public object? Value { get; }

    /// <summary>.NET type name for <see cref="Value"/> when available.</summary>
    public string? ValueType { get; }

    /// <summary>Underlying OfficeIMO custom-property type, when applicable.</summary>
    public string? CustomPropertyType { get; }
}
