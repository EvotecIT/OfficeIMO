namespace OfficeIMO.Project;

/// <summary>A predecessor relationship. Lag can be working/elapsed duration or a percentage.</summary>
public sealed class ProjectDependency : ProjectObject {
    internal ProjectDependency(ProjectDocument document) : base(document) { }
    /// <summary>The predecessor task, or null for an unresolved external source relationship.</summary>
    public ProjectTask? Predecessor { get; internal set; }
    /// <summary>The successor task.</summary>
    public ProjectTask Successor { get; internal set; } = null!;
    internal int SourcePredecessorUid { get; set; }
    private ProjectDependencyType? _type;
    private ProjectDuration? _lag;
    private decimal? _lagPercent;
    private bool _lagPercentIsElapsed, _lagPercentIsEstimated;
    private bool? _crossProject;
    private string? _crossProjectName;
    /// <summary>Dependency kind; absent source values remain absent.</summary>
    public ProjectDependencyType? Type { get => _type; set => Set(ref _type, value, true); }
    /// <summary>Working/elapsed lag; cannot be combined with a percentage lag.</summary>
    public ProjectDuration? Lag { get => _lag; set { if (value != null && _lagPercent != null) throw new InvalidOperationException("Clear percentage lag before setting duration lag."); Set(ref _lag, value, true); } }
    /// <summary>Percentage lag; cannot be combined with duration lag.</summary>
    public decimal? LagPercent {
        get => _lagPercent;
        set {
            if (value != null && _lag != null) throw new InvalidOperationException("Clear duration lag before setting percentage lag.");
            Set(ref _lagPercent, value, true);
            if (!value.HasValue) { Set(ref _lagPercentIsElapsed, false, true); Set(ref _lagPercentIsEstimated, false, true); }
        }
    }
    /// <summary>Whether percentage lag advances elapsed time instead of the successor's working calendar. Requires percentage lag.</summary>
    public bool LagPercentIsElapsed { get => _lagPercentIsElapsed; set { RequirePercentage(value); Set(ref _lagPercentIsElapsed, value, true); } }
    /// <summary>Whether percentage lag carries the source's estimated marker. Requires percentage lag; does not change arithmetic.</summary>
    public bool LagPercentIsEstimated { get => _lagPercentIsEstimated; set { RequirePercentage(value); Set(ref _lagPercentIsEstimated, value, true); } }
    internal int PercentageLagFormat => 19 + (LagPercentIsElapsed ? 1 : 0) + (LagPercentIsEstimated ? 32 : 0);
    private void RequirePercentage(bool value) { if (value && !LagPercent.HasValue) throw new InvalidOperationException("Set percentage lag before its elapsed or estimated flags."); }
    /// <summary>Whether the source declares a cross-project relationship.</summary>
    public bool? CrossProject { get => _crossProject; set => Set(ref _crossProject, value, true); }
    /// <summary>External project reference, preserved without opening it.</summary>
    public string? CrossProjectName { get => _crossProjectName; set => Set(ref _crossProjectName, value, true); }
}
