using System.Globalization;

namespace OfficeIMO.Project;

/// <summary>A calendar bucket with an exclusive finish.</summary>
public sealed class ProjectViewBucket {
    internal ProjectViewBucket(DateTime start, DateTime finish) { Start = start; Finish = finish; }
    /// <summary>Inclusive start.</summary>
    public DateTime Start { get; }
    /// <summary>Exclusive finish.</summary>
    public DateTime Finish { get; }
}

/// <summary>Immutable report row. Work excludes material and cost resources.</summary>
public sealed class ProjectViewRow {
    internal ProjectViewRow(int uid, string name, string group, DateTime? start, DateTime? finish,
        decimal? workHours, decimal? cost, int? complete, bool critical, bool summary,
        DateTime? baselineStart, DateTime? baselineFinish, decimal[] buckets) {
        Uid = uid; Name = name; Group = group; Start = start; Finish = finish; WorkHours = workHours;
        Cost = cost; PercentComplete = complete; IsCritical = critical; IsSummary = summary;
        BaselineStart = baselineStart; BaselineFinish = baselineFinish; BucketWorkHours = Array.AsReadOnly(buckets);
    }
    /// <summary>Task UID, or resource UID in resource layouts.</summary>
    public int Uid { get; }
    /// <summary>Entity name copied at report creation.</summary>
    public string Name { get; }
    /// <summary>Optional group label.</summary>
    public string Group { get; }
    /// <summary>Calculated start.</summary>
    public DateTime? Start { get; }
    /// <summary>Calculated finish.</summary>
    public DateTime? Finish { get; }
    /// <summary>Total work hours for the row, independent of the visible time window.</summary>
    public decimal? WorkHours { get; }
    /// <summary>Cost in project currency; null indicates incomplete data.</summary>
    public decimal? Cost { get; }
    /// <summary>Duration completion percentage; absent for resource rows.</summary>
    public int? PercentComplete { get; }
    /// <summary>Calculated critical status.</summary>
    public bool IsCritical { get; }
    /// <summary>Whether the row represents a summary task.</summary>
    public bool IsSummary { get; }
    /// <summary>Selected baseline start when available.</summary>
    public DateTime? BaselineStart { get; }
    /// <summary>Selected baseline finish when available.</summary>
    public DateTime? BaselineFinish { get; }
    /// <summary>Work hours intersecting each visible bucket, including actual and overtime work.</summary>
    public IReadOnlyList<decimal> BucketWorkHours { get; }
    /// <summary>Invariant text for an explicitly selected report column.</summary>
    public string GetText(ProjectViewColumn column) => column switch {
        ProjectViewColumn.Uid => Uid.ToString(CultureInfo.InvariantCulture),
        ProjectViewColumn.Name => Name,
        ProjectViewColumn.Start => Start?.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) ?? "",
        ProjectViewColumn.Finish => Finish?.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) ?? "",
        ProjectViewColumn.WorkHours => WorkHours?.ToString("0.##", CultureInfo.InvariantCulture) ?? "",
        ProjectViewColumn.Cost => Cost?.ToString("0.00", CultureInfo.InvariantCulture) ?? "",
        ProjectViewColumn.PercentComplete => PercentComplete?.ToString(CultureInfo.InvariantCulture) ?? "",
        ProjectViewColumn.Critical => IsCritical ? "Yes" : "No",
        _ => throw new ArgumentOutOfRangeException(nameof(column))
    };
}

/// <summary>A local dependency between included task rows.</summary>
public sealed class ProjectViewLink {
    internal ProjectViewLink(int predecessor, int successor, ProjectDependencyType type, ProjectDuration? lag, decimal? lagPercent) {
        PredecessorUid = predecessor; SuccessorUid = successor; Type = type; Lag = lag; LagPercent = lagPercent;
    }
    /// <summary>Predecessor UID.</summary>
    public int PredecessorUid { get; }
    /// <summary>Successor UID.</summary>
    public int SuccessorUid { get; }
    /// <summary>Dependency type.</summary>
    public ProjectDependencyType Type { get; }
    /// <summary>Working or elapsed duration lag, retaining its original units.</summary>
    public ProjectDuration? Lag { get; }
    /// <summary>Percentage lag, mutually exclusive with duration lag.</summary>
    public decimal? LagPercent { get; }
    /// <summary>Invariant lag label with explicit units.</summary>
    public string LagText => LagPercent.HasValue ? LagPercent.Value.ToString(CultureInfo.InvariantCulture) + "%" : Lag?.ToString() ?? "0";
}

/// <summary>Portable immutable report data; rendering creates independent drawing pages.</summary>
public sealed partial class ProjectView {
    internal readonly ProjectViewOptions Layout;
    internal ProjectView(string title, long revision, ProjectViewOptions layout, ProjectViewColumn[] columns,
        ProjectViewRow[] rows, ProjectViewBucket[] buckets, ProjectViewLink[] links, ProjectDiagnostic[] diagnostics) {
        Title = title; ModelRevision = revision; Layout = layout; Kind = layout.Kind;
        Columns = Array.AsReadOnly(columns); Rows = Array.AsReadOnly(rows);
        Buckets = Array.AsReadOnly(buckets); Links = Array.AsReadOnly(links);
        Report = new ProjectReport(revision, diagnostics);
    }
    /// <summary>Report title.</summary>
    public string Title { get; }
    /// <summary>Source revision. The captured report remains usable after subsequent document edits.</summary>
    public long ModelRevision { get; }
    /// <summary>Selected layout.</summary>
    public ProjectViewKind Kind { get; }
    /// <summary>Captured drawing and native report page width in points.</summary>
    public double PageWidth => Layout.PageWidth;
    /// <summary>Captured drawing and native report page height in points.</summary>
    public double PageHeight => Layout.PageHeight;
    /// <summary>Captured page margin in points.</summary>
    public double PageMargin => Layout.Margin;
    /// <summary>Maximum drawing pages or native presentation slides.</summary>
    public int MaxPages => Layout.MaxPages;
    /// <summary>Presentation fidelity notices. Native saved views, filters and styles are not interpreted by portable report layouts.</summary>
    public ProjectReport Report { get; }
    /// <summary>Selected columns.</summary>
    public IReadOnlyList<ProjectViewColumn> Columns { get; }
    /// <summary>Captured rows.</summary>
    public IReadOnlyList<ProjectViewRow> Rows { get; }
    /// <summary>Visible time buckets.</summary>
    public IReadOnlyList<ProjectViewBucket> Buckets { get; }
    /// <summary>Dependencies whose endpoints are both selected.</summary>
    public IReadOnlyList<ProjectViewLink> Links { get; }
}
