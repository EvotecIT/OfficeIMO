namespace OfficeIMO.Project;

/// <summary>Portable report layouts independent of native Project view tables.</summary>
public enum ProjectViewKind {
    /// <summary>Task bars against calendar dates.</summary>
    Gantt,
    /// <summary>Task work hours by calendar bucket.</summary>
    TaskUsage,
    /// <summary>Resource work hours by calendar bucket.</summary>
    ResourceUsage,
    /// <summary>Resource work-hour bars by calendar bucket.</summary>
    ResourceHistogram,
    /// <summary>Task nodes with predecessor relationships.</summary>
    Network,
    /// <summary>Task dates with compact labels.</summary>
    Timeline,
    /// <summary>Selected task columns.</summary>
    Table
}
/// <summary>Calendar bucket boundaries; weeks begin on Monday.</summary>
public enum ProjectViewTimescale {
    /// <summary>Local calendar days.</summary>
    Day,
    /// <summary>Weeks beginning on Monday.</summary>
    Week,
    /// <summary>Calendar months.</summary>
    Month
}
/// <summary>Stable report columns, with invariant units and formatting.</summary>
public enum ProjectViewColumn {
    /// <summary>Stable entity identity.</summary>
    Uid,
    /// <summary>Entity name.</summary>
    Name,
    /// <summary>Calculated start date.</summary>
    Start,
    /// <summary>Calculated finish date.</summary>
    Finish,
    /// <summary>Work hours, excluding material consumption.</summary>
    WorkHours,
    /// <summary>Cost in project currency.</summary>
    Cost,
    /// <summary>Duration completion percentage.</summary>
    PercentComplete,
    /// <summary>Calculated critical status.</summary>
    Critical
}
/// <summary>Optional grouping for task reports.</summary>
public enum ProjectViewGrouping {
    /// <summary>Preserve outline order.</summary>
    None,
    /// <summary>Group by immediate parent task.</summary>
    ParentTask
}

/// <summary>Selection and bounded layout settings captured when a report is created.</summary>
public sealed class ProjectViewOptions {
    /// <summary>Report layout.</summary>
    public ProjectViewKind Kind { get; set; } = ProjectViewKind.Gantt;
    /// <summary>Bucket granularity.</summary>
    public ProjectViewTimescale Timescale { get; set; } = ProjectViewTimescale.Week;
    /// <summary>Task identities to include; null selects all tasks.</summary>
    public IReadOnlyCollection<int>? TaskUids { get; set; }
    /// <summary>Resource identities to include; null selects all. Task layouts retain matching tasks and their summaries, using selected assignment work/cost totals (excluding task fixed costs), while dates and progress remain task-wide.</summary>
    public IReadOnlyCollection<int>? ResourceUids { get; set; }
    /// <summary>Report columns. Null uses layout-specific defaults.</summary>
    public IReadOnlyList<ProjectViewColumn>? Columns { get; set; }
    /// <summary>Include summary task rows.</summary>
    public bool IncludeSummaries { get; set; } = true;
    /// <summary>Restrict task rows and their assignments to critical tasks.</summary>
    public bool CriticalOnly { get; set; }
    /// <summary>Optional case-insensitive name substring filter.</summary>
    public string? NameContains { get; set; }
    /// <summary>Task grouping.</summary>
    public ProjectViewGrouping Grouping { get; set; }
    /// <summary>Optional baseline number from zero to ten.</summary>
    public int? BaselineNumber { get; set; }
    /// <summary>Inclusive visible boundary; omitted uses the selected schedule.</summary>
    public DateTime? Start { get; set; }
    /// <summary>Exclusive visible boundary; omitted uses the selected schedule.</summary>
    public DateTime? Finish { get; set; }
    /// <summary>Page width in points.</summary>
    public double PageWidth { get; set; } = 842;
    /// <summary>Page height in points.</summary>
    public double PageHeight { get; set; } = 595;
    /// <summary>Page margin in points.</summary>
    public double Margin { get; set; } = 28;
    /// <summary>Include a legend describing color and units.</summary>
    public bool ShowLegend { get; set; } = true;
    /// <summary>Maximum selected rows.</summary>
    public int MaxRows { get; set; } = 10000;
    /// <summary>Maximum time buckets before layout pagination.</summary>
    public int MaxBuckets { get; set; } = 1000;
    /// <summary>Maximum materialized row/bucket cells.</summary>
    public int MaxCells { get; set; } = 1000000;
    /// <summary>Maximum interval visits and bucket intersections across all rows, including summary aggregation.</summary>
    public int MaxIntervalVisits { get; set; } = 2000000;
    /// <summary>Maximum generated pages.</summary>
    public int MaxPages { get; set; } = 1000;
}
