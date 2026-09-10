namespace OfficeIMO.Project;

/// <summary>A baseline snapshot; stored values do not change when current task or assignment values change.</summary>
public sealed class ProjectBaseline : ProjectObject {
    internal ProjectBaseline(ProjectDocument document) : base(document) {
        TimephasedData = new ProjectCollection<ProjectTimephasedValue>(document, () => new ProjectTimephasedValue(document), owner: this);
    }
    /// <summary>Compact source timephased intervals.</summary>
    public ProjectCollection<ProjectTimephasedValue> TimephasedData { get; }

    private int? _number;
    /// <summary>Stored number; null represents an absent source value.</summary>
    public int? Number { get => _number; set => Set(ref _number, value); }

    private DateTime? _start;
    /// <summary>Stored local project date; null preserves an absent value.</summary>
    public DateTime? Start { get => _start; set => Set(ref _start, value); }

    private DateTime? _finish;
    /// <summary>Stored local project date; null preserves an absent value.</summary>
    public DateTime? Finish { get => _finish; set => Set(ref _finish, value); }

    private ProjectDuration? _duration;
    /// <summary>Stored duration; null represents an absent source value.</summary>
    public ProjectDuration? Duration { get => _duration; set => Set(ref _duration, value); }

    private ProjectWork? _work;
    /// <summary>Stored work; null represents an absent source value.</summary>
    public ProjectWork? Work { get => _work; set => Set(ref _work, value); }

    private decimal? _cost;
    /// <summary>Stored cost; null represents an absent source value.</summary>
    public decimal? Cost { get => _cost; set => Set(ref _cost, value); }

    private decimal? _fixedCost;
    /// <summary>Task baseline fixed cost. Not supported on resource or assignment baselines.</summary>
    public decimal? FixedCost { get => _fixedCost; set => Set(ref _fixedCost, value); }

    private decimal? _bcws;
    /// <summary>Stored bcws; null represents an absent source value.</summary>
    public decimal? Bcws { get => _bcws; set => Set(ref _bcws, value); }

    private decimal? _bcwp;
    /// <summary>Stored bcwp; null represents an absent source value.</summary>
    public decimal? Bcwp { get => _bcwp; set => Set(ref _bcwp, value); }
}
