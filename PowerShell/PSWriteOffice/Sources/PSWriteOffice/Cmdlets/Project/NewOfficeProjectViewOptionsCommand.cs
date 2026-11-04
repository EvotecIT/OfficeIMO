using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;
using OfficeIMO.Project;
namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Creates selection, timescale and page settings for Get-OfficeProjectView -Options.</summary>
/// <example><summary>Use the native OfficeProjectViewOptions command.</summary><code>$options = New-OfficeProjectViewOptions -Timescale Month
/// Get-OfficeProjectView -Document $project -Kind ResourceUsage -Options $options</code></example>
[Cmdlet("New", "OfficeProjectViewOptions", DefaultParameterSetName = "Options")]
[OutputType(typeof(ProjectViewOptions))]
public sealed class NewOfficeProjectViewOptionsCommand : OfficeCmdlet {
    /// <summary>Day, Week (default) or Month buckets; weeks begin on Monday.</summary>
    [Parameter(ParameterSetName = "Options")]
    public ProjectViewTimescale? Timescale { get; set; }
    /// <summary>Stable native task UIDs. Pass the Uid values of selected model objects; these are identities, not collection positions.</summary>
    [Parameter(ParameterSetName = "Options")]
    public int[]? Task { get; set; }
    /// <summary>Stable native resource UIDs. Pass the Uid values of selected model objects; these are identities, not collection positions.</summary>
    [Parameter(ParameterSetName = "Options")]
    public int[]? Resource { get; set; }
    /// <summary>Uid, Name, Start, Finish, WorkHours, Cost, PercentComplete, Critical; default depends on the layout.</summary>
    [Parameter(ParameterSetName = "Options")]
    public ProjectViewColumn[]? Column { get; set; }
    /// <summary>Case-insensitive task name filter.</summary>
    [Parameter(ParameterSetName = "Options")]
    public string? NameContains { get; set; }
    /// <summary>CriticalOnly for this operation.</summary>
    [Parameter(ParameterSetName = "Options")]
    public SwitchParameter CriticalOnly { get; set; }
    /// <summary>Omit summary task rows (shown by default).</summary>
    [Parameter(ParameterSetName = "Options")]
    public SwitchParameter NoSummaries { get; set; }
    /// <summary>None (outline order, default) or ParentTask.</summary>
    [Parameter(ParameterSetName = "Options")]
    public ProjectViewGrouping? Grouping { get; set; }
    /// <summary>Network view: Dependency (default) or Compact.</summary>
    [Parameter(ParameterSetName = "Options")]
    public ProjectNetworkLayout? NetworkLayout { get; set; }
    /// <summary>Status-date marker in Gantt and Timeline views.</summary>
    [Parameter(ParameterSetName = "Options")]
    public DateTime? StatusDate { get; set; }
    /// <summary>Baseline number 0-10 to compare against.</summary>
    [Parameter(ParameterSetName = "Options")]
    public int? Baseline { get; set; }
    /// <summary>First visible date (inclusive); default from the schedule.</summary>
    [Parameter(ParameterSetName = "Options")]
    public DateTime? Start { get; set; }
    /// <summary>Last visible date (exclusive); default from the schedule.</summary>
    [Parameter(ParameterSetName = "Options")]
    public DateTime? Finish { get; set; }
    /// <summary>Do not fill task bars with completion (shown by default).</summary>
    [Parameter(ParameterSetName = "Options")]
    public SwitchParameter NoProgress { get; set; }
    /// <summary>NoLegend for this operation.</summary>
    [Parameter(ParameterSetName = "Options")]
    public SwitchParameter NoLegend { get; set; }
    /// <summary>Default A4 landscape (842pt).</summary>
    [Parameter(ParameterSetName = "Options")]
    [OfficeLengthTransformation]
    public OfficeLength? PageWidth { get; set; }
    /// <summary>Default 595pt; the maximum height when pages fit their content.</summary>
    [Parameter(ParameterSetName = "Options")]
    [OfficeLengthTransformation]
    public OfficeLength? PageHeight { get; set; }
    /// <summary>Default 28pt.</summary>
    [Parameter(ParameterSetName = "Options")]
    [OfficeLengthTransformation]
    public OfficeLength? Margin { get; set; }
    /// <summary>Keep the full -PageHeight instead of trimming unused space (trimmed by default).</summary>
    [Parameter(ParameterSetName = "Options")]
    public SwitchParameter NoFitToContent { get; set; }
    /// <summary>Default 10000.</summary>
    [Parameter(ParameterSetName = "Options")]
    public int? MaximumRows { get; set; }
    /// <summary>Time buckets before pagination; default 1000.</summary>
    [Parameter(ParameterSetName = "Options")]
    public int? MaximumBuckets { get; set; }
    /// <summary>Default 1000000.</summary>
    [Parameter(ParameterSetName = "Options")]
    public int? MaximumCells { get; set; }
    /// <summary>Default 2000000.</summary>
    [Parameter(ParameterSetName = "Options")]
    public int? MaximumIntervalVisits { get; set; }
    /// <summary>Default 1000.</summary>
    [Parameter(ParameterSetName = "Options")]
    public int? MaximumPages { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new ProjectViewOptions();
        if (Timescale.HasValue) options.Timescale = Timescale.Value;
        if (Task != null) options.TaskUids = Task;
        if (Resource != null) options.ResourceUids = Resource;
        if (Column != null) options.Columns = Column;
        if (NameContains != null) options.NameContains = NameContains;
        if (OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(CriticalOnly)) is bool valueCriticalOnly) options.CriticalOnly = valueCriticalOnly;
        if (OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(NoSummaries)) is bool disabledNoSummaries) options.IncludeSummaries = !disabledNoSummaries;
        if (Grouping.HasValue) options.Grouping = Grouping.Value;
        if (NetworkLayout.HasValue) options.NetworkLayout = NetworkLayout.Value;
        if (StatusDate.HasValue) options.StatusDate = StatusDate.Value;
        if (Baseline.HasValue) options.BaselineNumber = Baseline.Value;
        if (Start.HasValue) options.Start = Start.Value;
        if (Finish.HasValue) options.Finish = Finish.Value;
        if (OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(NoProgress)) is bool disabledNoProgress) options.ShowProgress = !disabledNoProgress;
        if (OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(NoLegend)) is bool disabledNoLegend) options.ShowLegend = !disabledNoLegend;
        if (PageWidth.HasValue) options.PageWidth = PageWidth.Value.ToPoints();
        if (PageHeight.HasValue) options.PageHeight = PageHeight.Value.ToPoints();
        if (Margin.HasValue) options.Margin = Margin.Value.ToPoints();
        if (OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(NoFitToContent)) is bool disabledNoFitToContent) options.FitPageHeightToContent = !disabledNoFitToContent;
        if (MaximumRows.HasValue) options.MaxRows = MaximumRows.Value;
        if (MaximumBuckets.HasValue) options.MaxBuckets = MaximumBuckets.Value;
        if (MaximumCells.HasValue) options.MaxCells = MaximumCells.Value;
        if (MaximumIntervalVisits.HasValue) options.MaxIntervalVisits = MaximumIntervalVisits.Value;
        if (MaximumPages.HasValue) options.MaxPages = MaximumPages.Value;
        WriteObject(options);
    }
}
