namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>Captures a portable report from a current schedule without changing stored dates or native presentation data.</summary>
    public ProjectView CreateView(ProjectScheduleResult schedule, ProjectViewOptions? options = null, CancellationToken cancellationToken = default) {
        EnsureNotDisposed();
        if (schedule == null) throw new ArgumentNullException(nameof(schedule));
        CheckViewSchedule(schedule);
        var input = options ?? new ProjectViewOptions();
        var layout = ProjectViewBuilder.CopyOptions(input);
        var taskSelection = input.TaskUids; var resourceSelection = input.ResourceUids; var columnSelection = input.Columns;
        var taskIds = ProjectViewBuilder.SelectIds(taskSelection, TaskIndex.Keys, layout.MaxRows, nameof(input.TaskUids));
        var resourceIds = ProjectViewBuilder.SelectIds(resourceSelection, ResourceIndex.Keys, layout.MaxRows, nameof(input.ResourceUids));
        var columns = (columnSelection ?? new[] { ProjectViewColumn.Uid, ProjectViewColumn.Name, ProjectViewColumn.Start, ProjectViewColumn.Finish }).Take(17).ToArray();
        if (columns.Length == 0 || columns.Length > 16 || columns.Distinct().Count() != columns.Length || columns.Any(c => !Enum.IsDefined(typeof(ProjectViewColumn), c)))
            throw new ArgumentException("Select between one and sixteen distinct valid columns.", nameof(options));
        CheckViewSchedule(schedule); // Caller-owned collections can execute code while enumerated.
        bool resourceView = layout.Kind == ProjectViewKind.ResourceUsage || layout.Kind == ProjectViewKind.ResourceHistogram;
        if (resourceView) resourceIds = new HashSet<int>(Resources.Where(r =>
            (resourceIds == null || resourceIds.Contains(r.Uid)) && ProjectViewBuilder.Matches(r.Name, layout.NameContains)).Select(r => r.Uid));
        bool usage = resourceView || layout.Kind == ProjectViewKind.TaskUsage;
        if ((usage || resourceIds != null) && !schedule.CalculatedAssignments) throw new ArgumentException("Usage and resource-filtered reports require CalculateAssignments.", nameof(schedule));
        HashSet<int>? resourceTaskIds = null;
        if (!resourceView && resourceIds != null) {
            resourceTaskIds = new HashSet<int>();
            foreach (var assignment in schedule.Assignments.Where(a => resourceIds.Contains(a.ResourceUid))) {
                for (var task = TaskIndex[assignment.TaskUid]; task != null; task = task.Parent) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!resourceTaskIds.Add(task.Uid)) break;
                }
            }
        }
        var selected = schedule.Tasks.Where(t => (taskIds == null || taskIds.Contains(t.TaskUid))
            && (resourceTaskIds == null || resourceTaskIds.Contains(t.TaskUid))
            && (layout.IncludeSummaries || !t.IsSummary) && (!layout.CriticalOnly || t.IsCritical)
            && (resourceView || ProjectViewBuilder.Matches(TaskIndex[t.TaskUid].Name, layout.NameContains))).ToArray();
        if (selected.Length > layout.MaxRows) throw new InvalidOperationException("The selected task count exceeds MaxRows.");
        var selectedIds = new HashSet<int>(selected.Select(t => t.TaskUid));
        var assignments = schedule.Assignments.Where(a => (!resourceView || selectedIds.Contains(a.TaskUid)) && (resourceIds == null || resourceIds.Contains(a.ResourceUid))).ToArray();
        var buckets = ProjectViewBuilder.Buckets(layout, resourceView
            ? assignments.Select(a => (a.Start, a.Finish)).ToArray() : selected.Select(t => (t.Start, t.Finish)).ToArray());
        var rows = new List<ProjectViewRow>(); long intervalVisits = 0;
        if (resourceView) {
            var byResource = assignments.ToLookup(a => a.ResourceUid);
            foreach (var resource in Resources) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!resourceIds!.Contains(resource.Uid)) continue;
                var allocations = byResource[resource.Uid].ToArray();
                bool labor = (resource.Type ?? ProjectResourceType.Work) == ProjectResourceType.Work;
                ProjectViewBuilder.CheckCells(rows.Count + 1, buckets.Length, layout);
                rows.Add(new ProjectViewRow(resource.Uid, resource.Name ?? "", "", allocations.Length == 0 ? null : allocations.Min(a => a.Start),
                    allocations.Length == 0 ? null : allocations.Max(a => a.Finish), labor ? allocations.Sum(a => a.Work.Minutes) / 60m : (decimal?)null,
                    allocations.Any(a => !a.Cost.HasValue) ? null : allocations.Sum(a => a.Cost), null, false, false, null, null,
                    ProjectViewBuilder.WorkBuckets(labor ? allocations : Array.Empty<ProjectAssignmentSchedule>(), buckets, cancellationToken, ref intervalVisits, layout.MaxIntervalVisits)));
            }
        } else {
            var contributions = new Dictionary<int, List<ProjectAssignmentSchedule>>(); int contributionCount = 0;
            foreach (var allocation in assignments) {
                for (var owner = TaskIndex[allocation.TaskUid]; owner != null; owner = owner.Parent) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!selectedIds.Contains(owner.Uid)) continue;
                    if (++contributionCount > layout.MaxCells) throw new InvalidOperationException("Report assignment aggregation exceeds MaxCells.");
                    if (!contributions.TryGetValue(owner.Uid, out var values)) contributions.Add(owner.Uid, values = new List<ProjectAssignmentSchedule>());
                    values.Add(allocation);
                }
            }
            foreach (var result in selected) {
                cancellationToken.ThrowIfCancellationRequested();
                var task = TaskIndex[result.TaskUid];
                var baseline = layout.BaselineNumber.HasValue ? task.Baselines.FirstOrDefault(b => b.Number == layout.BaselineNumber) : null;
                ProjectViewBuilder.CheckCells(rows.Count + 1, buckets.Length, layout);
                // Summary buckets include descendant assignments even when their child rows are hidden.
                IEnumerable<ProjectAssignmentSchedule> rowAssignments = contributions.TryGetValue(result.TaskUid, out var values) ? values : Array.Empty<ProjectAssignmentSchedule>();
                var labor = rowAssignments.Where(a => (ResourceIndex[a.ResourceUid].Type ?? ProjectResourceType.Work) == ProjectResourceType.Work).ToArray();
                decimal? cost = resourceIds != null
                    ? rowAssignments.Any(a => !a.Cost.HasValue) ? null : rowAssignments.Sum(a => a.Cost)
                    : result.Calculation?.Cost ?? (schedule.CalculatedAssignments ? null : task.Cost);
                rows.Add(new ProjectViewRow(task.Uid, task.Name ?? "", layout.Grouping == ProjectViewGrouping.ParentTask ? task.Parent?.Name ?? "Root tasks" : "",
                    result.Start, result.Finish, resourceIds != null ? labor.Sum(a => a.Work.Minutes) / 60m : (result.Calculation?.Work ?? task.Work)?.Minutes / 60m,
                    cost, result.Calculation?.PercentComplete ?? task.PercentComplete,
                    result.IsCritical, result.IsSummary, baseline?.Start, baseline?.Finish,
                    ProjectViewBuilder.WorkBuckets(labor, buckets, cancellationToken, ref intervalVisits, layout.MaxIntervalVisits)));
            }
        }
        if (layout.Grouping != ProjectViewGrouping.None) rows = rows.OrderBy(r => r.Group, StringComparer.Ordinal).ToList();
        var included = new HashSet<int>(rows.Select(r => r.Uid));
        var links = resourceView ? Array.Empty<ProjectViewLink>() : Dependencies.Where(d => d.Predecessor != null && d.CrossProject != true
            && included.Contains(d.Predecessor.Uid) && included.Contains(d.Successor.Uid))
            .Select(d => new ProjectViewLink(d.Predecessor!.Uid, d.Successor.Uid, d.Type ?? ProjectDependencyType.FinishToStart, d.Lag, d.LagPercent)).ToArray();
        CheckViewSchedule(schedule); cancellationToken.ThrowIfCancellationRequested();
        var diagnostics = NativeSource == null ? Array.Empty<ProjectDiagnostic>() : new[] {
            new ProjectDiagnostic("PROJECT_VIEW_NATIVE_PRESENTATION", ProjectDiagnosticSeverity.Warning,
                "Native saved view tables, filters, formatting and graphical indicator tables are not interpreted. This report uses the explicit portable layout options.", "/Project/Views", true)
        };
        return new ProjectView(Title ?? Name ?? "Project report", Revision, layout, columns, rows.ToArray(), buckets, links, diagnostics);
    }

    private void CheckViewSchedule(ProjectScheduleResult schedule) {
        EnsureNotDisposed();
        if (HasActiveUpdate || schedule.Document != this || schedule.ModelRevision != Revision)
            throw new InvalidOperationException("Report creation requires a schedule for the current document revision outside an update scope.");
        schedule.Report.ThrowIfErrors();
        foreach (var source in schedule.ExternalSources) source.ValidateCurrent();
    }
}
