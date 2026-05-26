namespace OfficeIMO.Pdf;

/// <summary>
/// Simple AcroForm field information read from a PDF document.
/// </summary>
public sealed class PdfFormField {
    internal PdfFormField(int? objectNumber, string? name, string? partialName, string? fieldType, string? value, string? alternateName, string? mappingName, int? flags) {
        ObjectNumber = objectNumber;
        Name = name;
        PartialName = partialName;
        FieldType = fieldType;
        Value = value;
        AlternateName = alternateName;
        MappingName = mappingName;
        Flags = flags;
    }

    /// <summary>Indirect object number for the field dictionary, when known.</summary>
    public int? ObjectNumber { get; }

    /// <summary>Fully qualified field name when a name can be read.</summary>
    public string? Name { get; }

    /// <summary>Partial field name from the field dictionary.</summary>
    public string? PartialName { get; }

    /// <summary>Field type name, for example Tx, Btn, Ch, or Sig, when present or inherited.</summary>
    public string? FieldType { get; }

    /// <summary>Simple field value formatted for wrapper display, when present.</summary>
    public string? Value { get; }

    /// <summary>Alternate field name used as a user-facing label, when present.</summary>
    public string? AlternateName { get; }

    /// <summary>Mapping name used for export workflows, when present.</summary>
    public string? MappingName { get; }

    /// <summary>Raw field flags from /Ff, when present.</summary>
    public int? Flags { get; }
}
