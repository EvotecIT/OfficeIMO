namespace OfficeIMO.Project;

public sealed partial class ProjectTask {
    private DateTime? _earlyStart, _earlyFinish, _lateStart, _lateFinish;
    private decimal? _totalSlackMinutes, _freeSlackMinutes;
    /// <summary>Stored earliest start, independent of a new calculation result.</summary>
    public DateTime? EarlyStart { get => _earlyStart; set => Set(ref _earlyStart, value); }
    /// <summary>Stored earliest finish.</summary>
    public DateTime? EarlyFinish { get => _earlyFinish; set => Set(ref _earlyFinish, value); }
    /// <summary>Stored latest start.</summary>
    public DateTime? LateStart { get => _lateStart; set => Set(ref _lateStart, value); }
    /// <summary>Stored latest finish.</summary>
    public DateTime? LateFinish { get => _lateFinish; set => Set(ref _lateFinish, value); }
    /// <summary>Stored signed total float in working minutes.</summary>
    public decimal? TotalSlackMinutes { get => _totalSlackMinutes; set => Set(ref _totalSlackMinutes, value); }
    /// <summary>Stored signed free float in working minutes.</summary>
    public decimal? FreeSlackMinutes { get => _freeSlackMinutes; set => Set(ref _freeSlackMinutes, value); }
}
