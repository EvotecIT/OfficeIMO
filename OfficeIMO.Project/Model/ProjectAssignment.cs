namespace OfficeIMO.Project;

/// <summary>The relationship between a task and resource, with independently stored work, costs, and progress.</summary>
public sealed class ProjectAssignment : ProjectEntity {
    internal ProjectAssignment(ProjectDocument document, int uid) : base(document, uid) {
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
    /// <summary>The assigned task, or null for a preserved unresolved source reference.</summary>
    public ProjectTask? Task { get; internal set; }
    /// <summary>The resource, or null for an unassigned/preserved unresolved source reference.</summary>
    public ProjectResource? Resource { get; internal set; }
    internal int SourceTaskUid { get; set; }
    internal int SourceResourceUid { get; set; }

    private ProjectUnits? _units;
    /// <summary>Stored units; null represents an absent source value.</summary>
    public ProjectUnits? Units { get => _units; set => Set(ref _units, value, true); }

    private ProjectWork? _work;
    /// <summary>Stored work; null represents an absent source value.</summary>
    public ProjectWork? Work { get => _work; set => Set(ref _work, value, true); }

    private ProjectWork? _actualWork;
    /// <summary>Stored actual work; null represents an absent source value.</summary>
    public ProjectWork? ActualWork { get => _actualWork; set => Set(ref _actualWork, value, true); }

    private ProjectWork? _remainingWork;
    /// <summary>Stored remaining work; null represents an absent source value.</summary>
    public ProjectWork? RemainingWork { get => _remainingWork; set => Set(ref _remainingWork, value, true); }

    private ProjectWork? _overtimeWork;
    /// <summary>Stored overtime work; null represents an absent source value.</summary>
    public ProjectWork? OvertimeWork { get => _overtimeWork; set => Set(ref _overtimeWork, value, true); }

    private ProjectWork? _actualOvertimeWork;
    /// <summary>Stored actual overtime work; null represents an absent source value.</summary>
    public ProjectWork? ActualOvertimeWork { get => _actualOvertimeWork; set => Set(ref _actualOvertimeWork, value, true); }

    private DateTime? _start;
    /// <summary>Stored local project date; null preserves an absent value.</summary>
    public DateTime? Start { get => _start; set => Set(ref _start, value, true); }

    private DateTime? _finish;
    /// <summary>Stored local project date; null preserves an absent value.</summary>
    public DateTime? Finish { get => _finish; set => Set(ref _finish, value, true); }

    private DateTime? _actualStart;
    /// <summary>Stored local project date; null preserves an absent value.</summary>
    public DateTime? ActualStart { get => _actualStart; set => Set(ref _actualStart, value, true); }

    private DateTime? _actualFinish;
    /// <summary>Stored local project date; null preserves an absent value.</summary>
    public DateTime? ActualFinish { get => _actualFinish; set => Set(ref _actualFinish, value, true); }

    private decimal? _cost;
    /// <summary>Stored cost; null represents an absent source value.</summary>
    public decimal? Cost { get => _cost; set => Set(ref _cost, value, true); }

    private decimal? _actualCost;
    /// <summary>Stored actual cost; null represents an absent source value.</summary>
    public decimal? ActualCost { get => _actualCost; set => Set(ref _actualCost, value, true); }

    private decimal? _remainingCost;
    /// <summary>Stored remaining cost; null represents an absent source value.</summary>
    public decimal? RemainingCost { get => _remainingCost; set => Set(ref _remainingCost, value, true); }

    private int? _percentWorkComplete;
    /// <summary>Stored percent work complete; null represents an absent source value.</summary>
    public int? PercentWorkComplete { get => _percentWorkComplete; set => Set(ref _percentWorkComplete, value, true); }

    private string? _notes;
    /// <summary>Stored notes; null represents an absent source value.</summary>
    public string? Notes { get => _notes; set => Set(ref _notes, value); }
}
