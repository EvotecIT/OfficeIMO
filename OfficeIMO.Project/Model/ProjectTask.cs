namespace OfficeIMO.Project;

/// <summary>A task or summary in a project hierarchy; stored schedule values remain distinct from calculation.</summary>
public sealed partial class ProjectTask : ProjectNamedEntity {
    internal ProjectTask(ProjectDocument document, int uid) : base(document, uid) {
        TimephasedData = new ProjectCollection<ProjectTimephasedValue>(document, () => new ProjectTimephasedValue(document), owner: this);
        Baselines = new ProjectCollection<ProjectBaseline>(document, () => new ProjectBaseline(document), owner: this);
        CustomFields = new ProjectCollection<ProjectCustomFieldValue>(document, () => new ProjectCustomFieldValue(document), owner: this);
        Children = new ProjectTaskCollection(document, this);
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
    /// <summary>Immediate child tasks in outline order.</summary>
    public ProjectTaskCollection Children { get; }
    /// <summary>Parent summary, or null for a top-level task.</summary>
    public ProjectTask? Parent { get; internal set; }
    /// <summary>True for a summary task, including an imported summary with no modeled children.</summary>
    public bool IsSummary => Children.Count != 0 || SourceSummary;
    internal bool SourceSummary { get; set; }
    internal int? SourceOutlineLevel { get; set; }
    /// <summary>Moves this task and its children, retaining identities and relationships.</summary>
    public void MoveTo(ProjectTask? parent, int? index = null) => Document.MoveTask(this, parent, index);

    private int? _displayId;
    /// <summary>Display row identifier, separate from the stable UID.</summary>
    public int? DisplayId { get => _displayId; set => Set(ref _displayId, value); }

    private string? _wbs;
    /// <summary>Stored wbs; null represents an absent source value.</summary>
    public string? Wbs { get => _wbs; set => Set(ref _wbs, value); }

    private ProjectDuration? _duration;
    /// <summary>Stored duration; null represents an absent source value.</summary>
    public ProjectDuration? Duration { get => _duration; set => Set(ref _duration, value, true); }

    private ProjectDuration? _actualDuration;
    /// <summary>Stored actual duration; null represents an absent source value.</summary>
    public ProjectDuration? ActualDuration { get => _actualDuration; set => Set(ref _actualDuration, value, true); }

    private ProjectDuration? _remainingDuration;
    /// <summary>Stored remaining duration; null represents an absent source value.</summary>
    public ProjectDuration? RemainingDuration { get => _remainingDuration; set => Set(ref _remainingDuration, value, true); }

    private ProjectWork? _work;
    /// <summary>Stored work; null represents an absent source value.</summary>
    public ProjectWork? Work { get => _work; set => Set(ref _work, value, true); }

    private ProjectWork? _actualWork;
    /// <summary>Stored actual work; null represents an absent source value.</summary>
    public ProjectWork? ActualWork { get => _actualWork; set => Set(ref _actualWork, value, true); }

    private ProjectWork? _remainingWork;
    /// <summary>Stored remaining work; null represents an absent source value.</summary>
    public ProjectWork? RemainingWork { get => _remainingWork; set => Set(ref _remainingWork, value, true); }

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

    private DateTime? _deadline;
    /// <summary>Stored local project date; null preserves an absent value.</summary>
    public DateTime? Deadline { get => _deadline; set => Set(ref _deadline, value, true); }

    private DateTime? _constraintDate;
    /// <summary>Stored local project date; null preserves an absent value.</summary>
    public DateTime? ConstraintDate { get => _constraintDate; set => Set(ref _constraintDate, value, true); }

    private ProjectConstraintType? _constraintType;
    /// <summary>Stored constraint type; null represents an absent source value.</summary>
    public ProjectConstraintType? ConstraintType { get => _constraintType; set => Set(ref _constraintType, value, true); }

    private ProjectTaskType? _type;
    /// <summary>Stored type; null represents an absent source value.</summary>
    public ProjectTaskType? Type { get => _type; set => Set(ref _type, value, true); }

    private bool? _isManual;
    /// <summary>Stored is manual; null represents an absent source value.</summary>
    public bool? IsManual { get => _isManual; set => Set(ref _isManual, value, true); }

    private bool? _isMilestone;
    /// <summary>Stored is milestone; null represents an absent source value.</summary>
    public bool? IsMilestone { get => _isMilestone; set => Set(ref _isMilestone, value, true); }

    private bool? _effortDriven;
    /// <summary>Stored effort driven; null represents an absent source value.</summary>
    public bool? EffortDriven { get => _effortDriven; set => Set(ref _effortDriven, value, true); }

    private bool? _isActive;
    /// <summary>Stored is active; null represents an absent source value.</summary>
    public bool? IsActive { get => _isActive; set => Set(ref _isActive, value, true); }

    private bool? _isNull;
    /// <summary>Stored is null; null represents an absent source value.</summary>
    public bool? IsNull { get => _isNull; set => Set(ref _isNull, value, true); }

    private bool? _isCritical;
    /// <summary>Stored is critical; null represents an absent source value.</summary>
    public bool? IsCritical { get => _isCritical; set => Set(ref _isCritical, value); }

    private int? _percentComplete;
    /// <summary>Stored percent complete; null represents an absent source value.</summary>
    public int? PercentComplete { get => _percentComplete; set => Set(ref _percentComplete, value, true); }

    private int? _percentWorkComplete;
    /// <summary>Stored percent work complete; null represents an absent source value.</summary>
    public int? PercentWorkComplete { get => _percentWorkComplete; set => Set(ref _percentWorkComplete, value, true); }

    private int? _physicalPercentComplete;
    /// <summary>Stored physical percent complete; null represents an absent source value.</summary>
    public int? PhysicalPercentComplete { get => _physicalPercentComplete; set => Set(ref _physicalPercentComplete, value, true); }

    private int? _priority;
    /// <summary>Stored priority; null represents an absent source value.</summary>
    public int? Priority { get => _priority; set => Set(ref _priority, value, true); }

    private decimal? _cost;
    /// <summary>Stored cost; null represents an absent source value.</summary>
    public decimal? Cost { get => _cost; set => Set(ref _cost, value, true); }

    private decimal? _actualCost;
    /// <summary>Stored actual cost; null represents an absent source value.</summary>
    public decimal? ActualCost { get => _actualCost; set => Set(ref _actualCost, value, true); }

    private decimal? _remainingCost;
    /// <summary>Stored remaining cost; null represents an absent source value.</summary>
    public decimal? RemainingCost { get => _remainingCost; set => Set(ref _remainingCost, value, true); }

    private decimal? _fixedCost;
    /// <summary>Stored fixed cost; null represents an absent source value.</summary>
    public decimal? FixedCost { get => _fixedCost; set => Set(ref _fixedCost, value, true); }

    private string? _notes;
    /// <summary>Stored notes; null represents an absent source value.</summary>
    public string? Notes { get => _notes; set => Set(ref _notes, value); }

    private string? _contact;
    /// <summary>Stored contact; null represents an absent source value.</summary>
    public string? Contact { get => _contact; set => Set(ref _contact, value); }

    private string? _hyperlink;
    /// <summary>Stored hyperlink; null represents an absent source value.</summary>
    public string? Hyperlink { get => _hyperlink; set => Set(ref _hyperlink, value); }

    private string? _hyperlinkAddress;
    /// <summary>Stored hyperlink address; null represents an absent source value.</summary>
    public string? HyperlinkAddress { get => _hyperlinkAddress; set => Set(ref _hyperlinkAddress, value); }
}
