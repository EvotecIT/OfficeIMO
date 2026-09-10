namespace OfficeIMO.Project;

/// <summary>A custom-field value or lookup reference; its field definition retains the type.</summary>
public sealed class ProjectCustomFieldValue : ProjectObject {
    internal ProjectCustomFieldValue(ProjectDocument document) : base(document) {
    }

    private string? _fieldId;
    /// <summary>Stored field id; null represents an absent source value.</summary>
    public string? FieldId { get => _fieldId; set => Set(ref _fieldId, value); }

    private string? _value;
    /// <summary>Stored lexical value interpreted by the containing field or interval type.</summary>
    public string? Value { get => _value; set => Set(ref _value, value, true); }

    private string? _valueId;
    /// <summary>Stored value id; null represents an absent source value.</summary>
    public string? ValueId { get => _valueId; set => Set(ref _valueId, value); }

    private string? _valueGuid;
    /// <summary>Stored value guid; null represents an absent source value.</summary>
    public string? ValueGuid { get => _valueGuid; set => Set(ref _valueGuid, value); }

    private int? _durationFormat;
    /// <summary>Stored duration format; null represents an absent source value.</summary>
    public int? DurationFormat { get => _durationFormat; set => Set(ref _durationFormat, value); }
}
