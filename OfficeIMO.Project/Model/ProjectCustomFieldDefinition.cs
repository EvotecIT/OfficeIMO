namespace OfficeIMO.Project;

/// <summary>A Microsoft Project custom-field definition identified by its format field ID.</summary>
public sealed class ProjectCustomFieldDefinition : ProjectObject {
    internal ProjectCustomFieldDefinition(ProjectDocument document) : base(document) {
        LookupValues = new ProjectCollection<ProjectLookupValue>(document, () => new ProjectLookupValue(document), owner: this);
    }
    /// <summary>Allowed or suggested lookup values.</summary>
    public ProjectCollection<ProjectLookupValue> LookupValues { get; }

    private string? _fieldId;
    /// <summary>Stored field id; null represents an absent source value.</summary>
    public string? FieldId { get => _fieldId; set => Set(ref _fieldId, value); }

    private string? _fieldName;
    /// <summary>Stored field name; null represents an absent source value.</summary>
    public string? FieldName { get => _fieldName; set => Set(ref _fieldName, value); }

    private string? _alias;
    /// <summary>Stored alias; null represents an absent source value.</summary>
    public string? Alias { get => _alias; set => Set(ref _alias, value); }

    private int? _fieldType;
    /// <summary>Stored field type; null represents an absent source value.</summary>
    public int? FieldType { get => _fieldType; set => Set(ref _fieldType, value); }

    private string? _formula;
    /// <summary>Stored formula; null represents an absent source value.</summary>
    public string? Formula { get => _formula; set => Set(ref _formula, value, true); }

    private int? _summaryCalculation;
    /// <summary>XML CalculationType for summary rows: 0 retains no calculation, 1 rolls up children, and 2 evaluates the formula.</summary>
    public int? SummaryCalculation { get => _summaryCalculation; set => Set(ref _summaryCalculation, value, true); }

    private int? _rollupType;
    /// <summary>Stored XML RollupType. Unsupported operators are retained and diagnosed when calculation is requested.</summary>
    public int? RollupType { get => _rollupType; set => Set(ref _rollupType, value, true); }

    private string? _guid;
    /// <summary>Stored guid; null represents an absent source value.</summary>
    public string? Guid { get => _guid; set => Set(ref _guid, value); }

    private string? _lookupTableGuid;
    /// <summary>Ltuid identifying the shared outline-code lookup table.</summary>
    public string? LookupTableGuid { get => _lookupTableGuid; set => Set(ref _lookupTableGuid, value); }

    private bool? _restrictValues;
    /// <summary>Stored restrict values; null represents an absent source value.</summary>
    public bool? RestrictValues { get => _restrictValues; set => Set(ref _restrictValues, value); }
}
