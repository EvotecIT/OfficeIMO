namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    /// <summary>Analyzes stored assignment totals and uniform work/material costs without changing dates, actuals, rates, or source caches.</summary>
    public ProjectAssignmentAnalysis AnalyzeAssignments(CancellationToken cancellationToken = default) {
        EnsureNotDisposed();
        if (_batchDepth != 0) throw new InvalidOperationException("Finish the update scope before revision-bound analysis.");
        long revision = Revision;
        var diagnostics = new List<ProjectDiagnostic>(); var estimates = new List<ProjectAssignmentEstimate>();
        void Warn(string code, string message, string location) => diagnostics.Add(new ProjectDiagnostic(code, ProjectDiagnosticSeverity.Warning, message, location));
        foreach (var assignment in Assignments) {
            cancellationToken.ThrowIfCancellationRequested();
            var task = assignment.Task; var resource = assignment.Resource;
            string location = "/Assignment[UID=" + assignment.Uid + "]";
            if (assignment.Work is ProjectWork work && assignment.ActualWork is ProjectWork actual && assignment.RemainingWork is ProjectWork remaining && work.Minutes != actual.Minutes + remaining.Minutes)
                Warn("PROJECT_WORK_BALANCE", "Stored work differs from actual plus remaining work.", location);
            if (assignment.Cost.HasValue && assignment.ActualCost.HasValue && assignment.RemainingCost.HasValue && assignment.Cost != assignment.ActualCost + assignment.RemainingCost)
                Warn("PROJECT_COST_BALANCE", "Stored cost differs from actual plus remaining cost.", location);
            decimal? minutes = assignment.Work?.Minutes, cost = null;
            bool hasVariableRates = resource?.Rates.Count > 0;
            int? table = (int?)assignment.CostRateTable;
            if (resource?.Type == ProjectResourceType.Cost) cost = assignment.Cost;
            else if (NativeSource != null) Warn("PROJECT_NATIVE_RATE_ESTIMATE_UNSUPPORTED", "Native rate tables and assignment profiles are retained but not interpreted. No rate-based cost estimate was inferred.", location);
            else if (hasVariableRates || table > 0) Warn("PROJECT_RATE_ESTIMATE_UNSUPPORTED", "Dated rates or a non-default rate table require interval-specific cost calculation.", location);
            else if (resource?.Type == ProjectResourceType.Work) {
                if (!minutes.HasValue && task?.Duration is ProjectDuration duration && !duration.IsElapsed && assignment.Units.HasValue)
                    minutes = ProjectWorkEquation.Work(duration.Value * ProjectXmlValue.MinutesPerUnit(duration.Unit, false, this), assignment.Units.Value).Minutes;
                decimal overtime = assignment.OvertimeWork?.Minutes ?? 0m;
                if (minutes.HasValue && overtime > minutes.Value) Warn("PROJECT_OVERTIME_BALANCE", "Stored overtime exceeds total work.", location);
                else if (minutes.HasValue && resource.StandardRate.HasValue && (overtime == 0 || resource.OvertimeRate.HasValue))
                    cost = ProjectWorkEquation.WorkCost(new ProjectWork(minutes.Value), new ProjectWork(overtime), resource.StandardRate.Value, resource.OvertimeRate ?? 0m,
                        (resource.CostPerUse ?? 0m) * (assignment.Units?.Value ?? resource.MaxUnits?.Value ?? 1m));
            } else if (resource?.Type == ProjectResourceType.Material) {
                if (assignment.HasFixedRateUnits == false) Warn("PROJECT_MATERIAL_RATE_ESTIMATE_UNSUPPORTED", "Variable material usage requires a calendar-specific consumption calculation.", location);
                else if (assignment.Units.HasValue && resource.StandardRate.HasValue)
                    cost = ProjectWorkEquation.MaterialCost(assignment.Units.Value.Value, resource.StandardRate.Value, resource.CostPerUse ?? 0m);
            } else Warn("PROJECT_ASSIGNMENT_RESOURCE", "The resource is unspecified or unresolved; no rate estimate was inferred.", location);
            estimates.Add(new ProjectAssignmentEstimate(assignment, minutes, cost));
        }
        var grouped = Assignments.Where(a => a.Resource != null).GroupBy(a => a.Resource!).ToDictionary(g => g.Key, g => g.ToArray());
        var resources = new List<ProjectResourceTotals>();
        foreach (var resource in Resources) {
            cancellationToken.ThrowIfCancellationRequested();
            var totals = new ProjectResourceTotals(resource, grouped.TryGetValue(resource, out var assignments) ? assignments : Array.Empty<ProjectAssignment>());
            resources.Add(totals);
            if ((totals.Cost.HasValue && resource.Cost.HasValue && totals.Cost != resource.Cost) ||
                (totals.ActualCost.HasValue && resource.ActualCost.HasValue && totals.ActualCost != resource.ActualCost) ||
                (totals.RemainingCost.HasValue && resource.RemainingCost.HasValue && totals.RemainingCost != resource.RemainingCost))
                Warn("PROJECT_RESOURCE_CACHED_TOTAL", "Stored resource costs differ from the sum of stored assignment costs. Both observations are retained.", "/Resource[UID=" + resource.Uid + "]");
        }
        if (Revision != revision) throw new InvalidOperationException("The document changed during assignment analysis.");
        cancellationToken.ThrowIfCancellationRequested();
        return new ProjectAssignmentAnalysis(revision, resources, estimates, diagnostics);
    }
}
