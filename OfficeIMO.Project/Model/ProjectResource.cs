namespace OfficeIMO.Project;

/// <summary>A work, material, or cost resource; imported rates and actuals are retained without calculation.</summary>
public sealed class ProjectResource : ProjectNamedEntity {
    internal ProjectResource(ProjectDocument document, int uid) : base(document, uid) {
        TimephasedData = new ProjectCollection<ProjectTimephasedValue>(document, () => new ProjectTimephasedValue(document), owner: this);
        Baselines = new ProjectCollection<ProjectBaseline>(document, () => new ProjectBaseline(document), owner: this);
        CustomFields = new ProjectCollection<ProjectCustomFieldValue>(document, () => new ProjectCustomFieldValue(document), owner: this);
    }
    /// <summary>Compact source timephased intervals.</summary>
    public ProjectCollection<ProjectTimephasedValue> TimephasedData { get; }
    /// <summary>Stored baseline snapshots.</summary>
    public ProjectCollection<ProjectBaseline> Baselines { get; }
    /// <summary>Custom values and lookup references.</summary>
    public ProjectCollection<ProjectCustomFieldValue> CustomFields { get; }
    private ProjectCalendar? _calendar;
    /// <summary>Explicit calendar reference; null means no explicit calendar on this object.</summary>
    public ProjectCalendar? Calendar { get => _calendar; set { CheckReference(value); Set(ref _calendar, value, true); if (!Document.Loading) SourceCalendarUid = null; } }
    internal int? SourceCalendarUid { get; set; }

    private int? _displayId;
    /// <summary>Display row identifier, separate from the stable UID.</summary>
    public int? DisplayId { get => _displayId; set => Set(ref _displayId, value); }

    private ProjectResourceType? _type;
    /// <summary>Stored type; null represents an absent source value.</summary>
    public ProjectResourceType? Type { get => _type; set => Set(ref _type, value, true); }

    private string? _initials;
    /// <summary>Stored initials; null represents an absent source value.</summary>
    public string? Initials { get => _initials; set => Set(ref _initials, value); }

    private string? _group;
    /// <summary>Stored group; null represents an absent source value.</summary>
    public string? Group { get => _group; set => Set(ref _group, value); }

    private string? _emailAddress;
    /// <summary>Stored email address; null represents an absent source value.</summary>
    public string? EmailAddress { get => _emailAddress; set => Set(ref _emailAddress, value); }

    private string? _materialLabel;
    /// <summary>Stored material label; null represents an absent source value.</summary>
    public string? MaterialLabel { get => _materialLabel; set => Set(ref _materialLabel, value); }

    private ProjectUnits? _maxUnits;
    /// <summary>Stored max units; null represents an absent source value.</summary>
    public ProjectUnits? MaxUnits { get => _maxUnits; set => Set(ref _maxUnits, value, true); }

    private decimal? _standardRate;
    /// <summary>Stored standard rate; null represents an absent source value.</summary>
    public decimal? StandardRate { get => _standardRate; set => Set(ref _standardRate, value, true); }

    private decimal? _overtimeRate;
    /// <summary>Stored overtime rate; null represents an absent source value.</summary>
    public decimal? OvertimeRate { get => _overtimeRate; set => Set(ref _overtimeRate, value, true); }

    private decimal? _costPerUse;
    /// <summary>Stored cost per use; null represents an absent source value.</summary>
    public decimal? CostPerUse { get => _costPerUse; set => Set(ref _costPerUse, value, true); }

    private decimal? _cost;
    /// <summary>Stored cost; null represents an absent source value.</summary>
    public decimal? Cost { get => _cost; set => Set(ref _cost, value, true); }

    private decimal? _actualCost;
    /// <summary>Stored actual cost; null represents an absent source value.</summary>
    public decimal? ActualCost { get => _actualCost; set => Set(ref _actualCost, value, true); }

    private ProjectWork? _work;
    /// <summary>Stored work; null represents an absent source value.</summary>
    public ProjectWork? Work { get => _work; set => Set(ref _work, value, true); }

    private ProjectWork? _actualWork;
    /// <summary>Stored actual work; null represents an absent source value.</summary>
    public ProjectWork? ActualWork { get => _actualWork; set => Set(ref _actualWork, value, true); }

    private ProjectWork? _remainingWork;
    /// <summary>Stored remaining work; null represents an absent source value.</summary>
    public ProjectWork? RemainingWork { get => _remainingWork; set => Set(ref _remainingWork, value, true); }

    private string? _notes;
    /// <summary>Stored notes; null represents an absent source value.</summary>
    public string? Notes { get => _notes; set => Set(ref _notes, value); }

    private bool? _isNull;
    /// <summary>Stored is null; null represents an absent source value.</summary>
    public bool? IsNull { get => _isNull; set => Set(ref _isNull, value, true); }
}
