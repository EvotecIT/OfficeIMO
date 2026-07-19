namespace OfficeIMO.Email;

/// <summary>Logical Outlook field types stored in a PropertyDefinition stream.</summary>
public enum OutlookUserPropertyType {
    /// <summary>The definition uses a type OfficeIMO does not currently interpret.</summary>
    Unknown = -1,
    /// <summary>Text field.</summary>
    Text = 0,
    /// <summary>Floating-point number field.</summary>
    Number = 1,
    /// <summary>Percentage field.</summary>
    Percent = 2,
    /// <summary>Fixed-point currency field.</summary>
    Currency = 3,
    /// <summary>Yes/no field.</summary>
    Boolean = 4,
    /// <summary>Date and time field.</summary>
    DateTime = 5,
    /// <summary>Duration field whose MAPI value is a number of minutes.</summary>
    Duration = 6,
    /// <summary>Combination field selecting the first non-empty component.</summary>
    Combination = 7,
    /// <summary>Calculated formula field.</summary>
    Formula = 8,
    /// <summary>Combination field concatenating its components.</summary>
    Concatenation = 12,
    /// <summary>Multiple keyword strings.</summary>
    Keywords = 13,
    /// <summary>32-bit integer field.</summary>
    Integer = 14
}

/// <summary>Parse state of an item's PidLidPropertyDefinitionStream value.</summary>
public enum OutlookUserPropertyDefinitionState {
    /// <summary>The item does not contain a PropertyDefinition stream.</summary>
    Missing,
    /// <summary>The complete stream was decoded.</summary>
    Valid,
    /// <summary>The stream declares a version newer than the documented V1/V2 formats.</summary>
    UnsupportedVersion,
    /// <summary>The stream is truncated or structurally invalid.</summary>
    Corrupt
}

/// <summary>One definition retained from an Outlook PropertyDefinition stream.</summary>
public sealed class OutlookUserPropertyDefinition {
    internal OutlookUserPropertyDefinition(string name, uint flags, ushort variantType, uint dispatchId,
        OutlookUserPropertyType fieldType, string formula, string validationRule, string validationText,
        bool isVersion2, byte[] rawDefinition) {
        Name = name;
        Flags = flags;
        VariantType = variantType;
        DispatchId = dispatchId;
        FieldType = fieldType;
        Formula = formula;
        ValidationRule = validationRule;
        ValidationText = validationText;
        IsVersion2 = isVersion2;
        RawDefinition = rawDefinition;
    }

    /// <summary>Field name used by the corresponding PS_PUBLIC_STRINGS named property.</summary>
    public string Name { get; }

    /// <summary>Raw PropertyDefinition flags.</summary>
    public uint Flags { get; }

    /// <summary>VARENUM value stored by Outlook.</summary>
    public ushort VariantType { get; }

    /// <summary>Dispatch identifier; custom fields normally use zero.</summary>
    public uint DispatchId { get; }

    /// <summary>Logical Outlook field type.</summary>
    public OutlookUserPropertyType FieldType { get; }

    /// <summary>ANSI calculation formula retained from the definition.</summary>
    public string Formula { get; }

    /// <summary>ANSI validation rule retained from the definition.</summary>
    public string ValidationRule { get; }

    /// <summary>ANSI validation failure text retained from the definition.</summary>
    public string ValidationText { get; }

    /// <summary>Whether this definition uses the PropDefV2 layout.</summary>
    public bool IsVersion2 { get; }

    /// <summary>Whether the definition represents a user-defined rather than built-in field.</summary>
    public bool IsCustom => (Flags & 0x00000001U) != 0;

    internal byte[] RawDefinition { get; }
}

/// <summary>A user-defined Outlook field joined with its PS_PUBLIC_STRINGS value.</summary>
public sealed class OutlookUserProperty {
    internal OutlookUserProperty(string name, MapiProperty? property,
        OutlookUserPropertyDefinition? definition) {
        Name = name;
        Property = property;
        Definition = definition;
    }

    /// <summary>Case-insensitive field name.</summary>
    public string Name { get; }

    /// <summary>Logical field type, or Unknown for a value without a definition.</summary>
    public OutlookUserPropertyType FieldType => Definition?.FieldType ?? OutlookUserPropertyType.Unknown;

    /// <summary>MAPI wire type of the value, or null when the definition has no stored value.</summary>
    public MapiPropertyType? WireType => Property?.PropertyType;

    /// <summary>Decoded value as retained by the MAPI property layer.</summary>
    public object? Value => Property?.Value;

    /// <summary>True when the item contains a matching PropertyDefinition entry.</summary>
    public bool HasDefinition => Definition != null;

    /// <summary>True when the item contains the corresponding named value property.</summary>
    public bool HasValue => Property != null;

    /// <summary>Decoded definition, when present and structurally valid.</summary>
    public OutlookUserPropertyDefinition? Definition { get; }

    internal MapiProperty? Property { get; }
}
