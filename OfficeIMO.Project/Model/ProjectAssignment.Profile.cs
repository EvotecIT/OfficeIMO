namespace OfficeIMO.Project;

public sealed partial class ProjectAssignment {
    private ProjectCostRateTable? _costRateTable;
    private decimal? _delayMinutes;
    private bool? _hasFixedRateUnits;
    private int? _materialRateScale;
    private ProjectWorkContour? _workContour;
    private DateTime? _stop, _resume;
    /// <summary>End of recorded actual work, if supplied by the source or caller.</summary>
    public DateTime? Stop { get => _stop; set => Set(ref _stop, value, true); }
    /// <summary>Earliest stored restart of remaining work after an interruption.</summary>
    public DateTime? Resume { get => _resume; set => Set(ref _resume, value, true); }
    /// <summary>Resource cost table selected for this assignment; null uses table A.</summary>
    public ProjectCostRateTable? CostRateTable { get => _costRateTable; set => Set(ref _costRateTable, value, true); }
    /// <summary>Working-minute delay from task start to assignment start.</summary>
    public decimal? DelayMinutes { get => _delayMinutes; set => Set(ref _delayMinutes, value, true); }
    /// <summary>True for fixed material quantities, false for consumption per time unit. Work-resource allocations normally use true.</summary>
    public bool? HasFixedRateUnits { get => _hasFixedRateUnits; set => Set(ref _hasFixedRateUnits, value, true); }
    /// <summary>Variable material consumption unit: 1 minute, 2 hour, 3 day, 4 week, 5 month. Units stores the quantity per selected unit.</summary>
    public int? MaterialRateScale { get => _materialRateScale; set => Set(ref _materialRateScale, value, true); }
    /// <summary>Named allocation profile, or a custom profile supplied by remaining-work timephased intervals.</summary>
    public ProjectWorkContour? WorkContour { get => _workContour; set => Set(ref _workContour, value, true); }
}

/// <summary>How work is distributed along an assignment's working duration.</summary>
public enum ProjectWorkContour {
    /// <summary>Uniform allocation.</summary>
    Flat = 0,
    /// <summary>More work near the end.</summary>
    BackLoaded = 1,
    /// <summary>More work near the beginning.</summary>
    FrontLoaded = 2,
    /// <summary>Two allocation peaks.</summary>
    DoublePeak = 3,
    /// <summary>An early allocation peak.</summary>
    EarlyPeak = 4,
    /// <summary>A late allocation peak.</summary>
    LatePeak = 5,
    /// <summary>A central allocation peak.</summary>
    Bell = 6,
    /// <summary>A broad central allocation plateau.</summary>
    Turtle = 7,
    /// <summary>Explicit timephased allocation, including gaps and splits.</summary>
    Custom = 8
}
