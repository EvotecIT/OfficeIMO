namespace OfficeIMO.Project;

/// <summary>A hierarchical lookup table and masks used by one or more outline-code fields.</summary>
public sealed class ProjectOutlineCodeDefinition : ProjectObject {
    internal ProjectOutlineCodeDefinition(ProjectDocument document) : base(document) {
        Masks = new ProjectCollection<ProjectOutlineCodeMask>(document, () => new ProjectOutlineCodeMask(document), owner: this);
        Values = new ProjectCollection<ProjectOutlineCodeLookupValue>(document, () => new ProjectOutlineCodeLookupValue(document), owner: this);
    }
    /// <summary>Ordered masks belonging to this definition.</summary>
    public ProjectCollection<ProjectOutlineCodeMask> Masks { get; }
    /// <summary>Ordered values belonging to this definition.</summary>
    public ProjectCollection<ProjectOutlineCodeLookupValue> Values { get; }
    private string? _guid;
    /// <summary>Lookup-table identity referenced by a custom field's LookupTableGuid.</summary>
    public string? Guid { get => _guid; set => Set(ref _guid, value); }
    private string? _fieldId;
    /// <summary>Optional legacy field identity.</summary>
    public string? FieldId { get => _fieldId; set => Set(ref _fieldId, value); }
    private string? _fieldName;
    /// <summary>Optional legacy field name.</summary>
    public string? FieldName { get => _fieldName; set => Set(ref _fieldName, value); }
    private string? _alias;
    /// <summary>Optional display alias.</summary>
    public string? Alias { get => _alias; set => Set(ref _alias, value); }
    private bool? _leafOnly;
    /// <summary>Whether selection is restricted to leaf values.</summary>
    public bool? LeafOnly { get => _leafOnly; set => Set(ref _leafOnly, value); }
    private bool? _allLevelsRequired;
    /// <summary>Whether selected values must contain every mask level.</summary>
    public bool? AllLevelsRequired { get => _allLevelsRequired; set => Set(ref _allLevelsRequired, value); }
    private bool? _onlyTableValuesAllowed;
    /// <summary>Whether free-text outline values are forbidden.</summary>
    public bool? OnlyTableValuesAllowed { get => _onlyTableValuesAllowed; set => Set(ref _onlyTableValuesAllowed, value); }
}
