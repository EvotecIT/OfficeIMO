namespace OfficeIMO.Project;

public sealed partial class ProjectTask {
    private bool? _isRecurring;
    private bool? _levelingCanSplit;
    /// <summary>Whether explicit resource leveling may interrupt remaining work. False permits delays only.</summary>
    public bool? LevelingCanSplit { get => _levelingCanSplit; set => Set(ref _levelingCanSplit, value, true); }
    /// <summary>Whether this stored task belongs to an expanded recurring series. This marker does not contain or regenerate the native recurrence rule.</summary>
    public bool? IsRecurring { get => _isRecurring; set => Set(ref _isRecurring, value, true); }
    private bool? _ignoreResourceCalendar;
    private ProjectCostAccrual? _fixedCostAccrual;
    private DateTime? _stop, _resume;
    private ProjectEarnedValueMethod? _earnedValueMethod;
    private ProjectDuration? _levelingDelay;
    /// <summary>Explicit delay from the dependency/constraint early-start anchor. Working and elapsed delays retain their different calendar semantics.</summary>
    public ProjectDuration? LevelingDelay { get => _levelingDelay; set => Set(ref _levelingDelay, value, true); }
    /// <summary>Use the explicit task calendar without intersecting resource calendars.</summary>
    public bool? IgnoreResourceCalendar { get => _ignoreResourceCalendar; set => Set(ref _ignoreResourceCalendar, value, true); }
    /// <summary>When the task's fixed cost accrues; null uses prorated accrual.</summary>
    public ProjectCostAccrual? FixedCostAccrual { get => _fixedCostAccrual; set => Set(ref _fixedCostAccrual, value, true); }
    /// <summary>End of the recorded completed portion, separate from task finish.</summary>
    public DateTime? Stop { get => _stop; set => Set(ref _stop, value, true); }
    /// <summary>Stored beginning of remaining work after a progress interruption.</summary>
    public DateTime? Resume { get => _resume; set => Set(ref _resume, value, true); }
    /// <summary>Completion measure used to calculate earned value.</summary>
    public ProjectEarnedValueMethod? EarnedValueMethod { get => _earnedValueMethod; set => Set(ref _earnedValueMethod, value, true); }
}

/// <summary>Which independently stored completion percentage earns baseline cost.</summary>
public enum ProjectEarnedValueMethod {
    /// <summary>Duration-based completion percentage.</summary>
    PercentComplete = 0,
    /// <summary>Physical completion supplied by the caller.</summary>
    PhysicalPercentComplete = 1
}
