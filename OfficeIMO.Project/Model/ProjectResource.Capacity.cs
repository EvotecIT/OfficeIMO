namespace OfficeIMO.Project;

public sealed partial class ProjectResource {
    /// <summary>Dated capacity periods. Outside explicitly listed periods capacity is zero; an empty collection uses MaxUnits.</summary>
    public ProjectCollection<ProjectResourceAvailability> AvailabilityPeriods { get; }
    /// <summary>Dated cost rates grouped by table A through E. Work rates are expressed per hour; material rates are per material unit.</summary>
    public ProjectCollection<ProjectResourceRate> Rates { get; }
    private ProjectCostAccrual? _accrueAt;
    /// <summary>When work or material costs accrue. Per-use costs are charged once when the assignment starts.</summary>
    public ProjectCostAccrual? AccrueAt { get => _accrueAt; set => Set(ref _accrueAt, value, true); }
    private bool? _canLevel;
    /// <summary>Whether explicit resource leveling may move this resource's assignments.</summary>
    public bool? CanLevel { get => _canLevel; set => Set(ref _canLevel, value, true); }
}

/// <summary>When a task's fixed cost or a resource's usage cost enters its timephased cost plan.</summary>
public enum ProjectCostAccrual {
    /// <summary>Charge at the beginning.</summary>
    Start = 1,
    /// <summary>Charge at completion.</summary>
    End = 2,
    /// <summary>Distribute according to the work or consumption performed.</summary>
    Prorated = 3
}

/// <summary>One of the five cost rate tables supported by Project interchange.</summary>
public enum ProjectCostRateTable {
    /// <summary>Default table A.</summary>
    A = 0,
    /// <summary>Alternative table B.</summary>
    B = 1,
    /// <summary>Alternative table C.</summary>
    C = 2,
    /// <summary>Alternative table D.</summary>
    D = 3,
    /// <summary>Alternative table E.</summary>
    E = 4
}

/// <summary>A resource's available allocation fraction during an inclusive, minute-resolution range.</summary>
public sealed class ProjectResourceAvailability : ProjectObject {
    internal ProjectResourceAvailability(ProjectDocument document) : base(document) { }
    private DateTime? _from, _through;
    private ProjectUnits? _units;
    /// <summary>First available local minute; null leaves the beginning unbounded.</summary>
    public DateTime? From { get => _from; set => Set(ref _from, value, true); }
    /// <summary>Last available local minute; null leaves the end unbounded.</summary>
    public DateTime? Through { get => _through; set => Set(ref _through, value, true); }
    /// <summary>Available fraction, where 1 means one full resource. Missing units use the resource's maximum units.</summary>
    public ProjectUnits? Units { get => _units; set => Set(ref _units, value, true); }
}

/// <summary>Hourly work rates or per-unit material prices for a half-open local time range in one rate table.</summary>
public sealed class ProjectResourceRate : ProjectObject {
    internal ProjectResourceRate(ProjectDocument document) : base(document) { }
    private DateTime? _from, _to;
    private ProjectCostRateTable? _table;
    private decimal? _standardRate, _overtimeRate, _costPerUse;
    private int? _standardRateFormat, _overtimeRateFormat;
    /// <summary>Inclusive effective date. Required when writing a dated rate to MSPDI.</summary>
    public DateTime? From { get => _from; set => Set(ref _from, value, true); }
    /// <summary>Exclusive effective end. Adjacent rate periods may share this boundary.</summary>
    public DateTime? To { get => _to; set => Set(ref _to, value, true); }
    /// <summary>Rate table; an absent table means A.</summary>
    public ProjectCostRateTable? Table { get => _table; set => Set(ref _table, value, true); }
    /// <summary>Currency units per work hour or material unit, independent of display format.</summary>
    public decimal? StandardRate { get => _standardRate; set => Set(ref _standardRate, value, true); }
    /// <summary>Currency units per overtime hour.</summary>
    public decimal? OvertimeRate { get => _overtimeRate; set => Set(ref _overtimeRate, value, true); }
    /// <summary>One charge per assignment in currency units.</summary>
    public decimal? CostPerUse { get => _costPerUse; set => Set(ref _costPerUse, value, true); }
    /// <summary>Source display-unit code; it does not rescale StandardRate.</summary>
    public int? StandardRateFormat { get => _standardRateFormat; set => Set(ref _standardRateFormat, value); }
    /// <summary>Source display-unit code; it does not rescale OvertimeRate.</summary>
    public int? OvertimeRateFormat { get => _overtimeRateFormat; set => Set(ref _overtimeRateFormat, value); }
}
