namespace OfficeIMO.Project;

/// <summary>One component of a hierarchical outline lookup value.</summary>
public sealed class ProjectOutlineCodeLookupValue : ProjectObject {
    internal ProjectOutlineCodeLookupValue(ProjectDocument document) : base(document) {
    }
    private int? _valueId;
    /// <summary>Identity within the lookup table.</summary>
    public int? ValueId { get => _valueId; set => Set(ref _valueId, value); }
    private string? _guid;
    /// <summary>Value identity referenced by task or resource ValueGUID.</summary>
    public string? Guid { get => _guid; set => Set(ref _guid, value); }
    private int? _parentValueId;
    /// <summary>Parent identity; zero or null denotes a root.</summary>
    public int? ParentValueId { get => _parentValueId; set => Set(ref _parentValueId, value); }
    private int? _type;
    /// <summary>Stored lookup value type; 21 denotes text.</summary>
    public int? Type { get => _type; set => Set(ref _type, value); }
    private string? _value;
    /// <summary>This level's value component, without ancestor components.</summary>
    public string? Value { get => _value; set => Set(ref _value, value); }
    private string? _description;
    /// <summary>Optional description shown in the lookup table.</summary>
    public string? Description { get => _description; set => Set(ref _description, value); }
}
