namespace OfficeIMO.Project;

/// <summary>An entry in a custom field's lookup list.</summary>
public sealed class ProjectLookupValue : ProjectObject {
    internal ProjectLookupValue(ProjectDocument document) : base(document) {
    }

    private int? _id;
    /// <summary>Stored id; null represents an absent source value.</summary>
    public int? Id { get => _id; set => Set(ref _id, value); }

    private string? _value;
    /// <summary>Stored lexical value interpreted by the containing field or interval type.</summary>
    public string? Value { get => _value; set => Set(ref _value, value); }

    private string? _description;
    /// <summary>Stored description; null represents an absent source value.</summary>
    public string? Description { get => _description; set => Set(ref _description, value); }

    private string? _guid;
    /// <summary>Stored guid; null represents an absent source value.</summary>
    public string? Guid { get => _guid; set => Set(ref _guid, value); }
}
