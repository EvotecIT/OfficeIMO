namespace OfficeIMO.Project;

internal sealed partial class ProjectTaskAllocation {
    private (decimal? Cost, decimal? Actual) CalculateCosts(Entry entry, List<ProjectAssignmentInterval> intervals, DateTime start, DateTime finish, List<ProjectCostInterval> charges) {
        var assignment = entry.Assignment; var resource = assignment.Resource!;
        if (resource.Type == ProjectResourceType.Cost) {
            decimal? value = assignment.Cost, actual = assignment.ActualCost ?? 0m;
            if (value.HasValue) { charges.Add(new ProjectCostInterval(start, start, actual ?? 0m, true)); charges.Add(new ProjectCostInterval(finish, finish, value.Value - (actual ?? 0m), false)); }
            return (value, actual);
        }
        var table = assignment.CostRateTable ?? ProjectCostRateTable.A;
        var usage = new List<ProjectCostInterval>(); bool complete = true;
        foreach (var interval in intervals) {
            _token.ThrowIfCancellationRequested();
            var points = entry.Resource.RateBoundaries(interval.Start, interval.Finish, table).Concat(new[] { interval.Start, interval.Finish }).Distinct().OrderBy(d => d).ToArray();
            if (points.Length == 1) {
                var rate = entry.Resource.Rate(interval.Start, table);
                decimal? pointRate = entry.IsFixedMaterial ? rate.Standard : rate.Overtime;
                decimal quantity = entry.IsFixedMaterial ? interval.Work.Minutes / 60m : interval.OvertimeWork.Minutes / 60m;
                if (!pointRate.HasValue) complete = false;
                else { usage.Add(new ProjectCostInterval(interval.Start, interval.Finish, quantity * pointRate.Value, interval.IsActual)); CheckCount(usage.Count); }
                continue;
            }
            for (int index = 1; index < points.Length; index++) {
                var rate = entry.Resource.Rate(points[index - 1], table);
                decimal ratio = (points[index].Ticks - points[index - 1].Ticks) / (decimal)(interval.Finish.Ticks - interval.Start.Ticks);
                if (!rate.Standard.HasValue || interval.OvertimeWork.Minutes > 0 && !rate.Overtime.HasValue) { complete = false; continue; }
                decimal amount = resource.Type == ProjectResourceType.Material ? interval.Work.Minutes / 60m * ratio * rate.Standard.Value :
                    ProjectWorkEquation.WorkCost(new ProjectWork(interval.Work.Minutes * ratio), new ProjectWork(interval.OvertimeWork.Minutes * ratio), rate.Standard.Value, rate.Overtime ?? 0m);
                usage.Add(new ProjectCostInterval(points[index - 1], points[index], amount, interval.IsActual)); CheckCount(usage.Count);
            }
        }
        if (!complete) {
            Warn("PROJECT_RATE_INCOMPLETE", "The selected rate table does not cover every work interval. Total cost was not inferred.", assignment);
            return (null, assignment.ActualCost);
        }
        var firstRate = entry.Resource.Rate(start, table); decimal perUse = firstRate.PerUse ?? 0m;
        if (resource.Type == ProjectResourceType.Work) perUse *= entry.Units;
        bool began = entry.Actual > 0 || assignment.ActualStart.HasValue;
        bool ended = entry.Remaining == 0 && (assignment.ActualFinish.HasValue || entry.Actual > 0);
        switch (resource.AccrueAt ?? ProjectCostAccrual.Prorated) {
            case ProjectCostAccrual.Start: charges.Add(new ProjectCostInterval(start, start, usage.Sum(c => c.Cost), began)); break;
            case ProjectCostAccrual.End: charges.Add(new ProjectCostInterval(finish, finish, usage.Sum(c => c.Cost), ended)); break;
            default: charges.AddRange(usage); break;
        }
        if (perUse != 0) charges.Add(new ProjectCostInterval(start, start, perUse, began));
        decimal total = charges.Sum(c => c.Cost), computedActual = charges.Where(c => c.IsActual).Sum(c => c.Cost);
        var stored = assignment.TimephasedData.Where(v => v.Type == 6).ToArray();
        if (!_options.RecalculateActualCosts && (assignment.ActualCost.HasValue || stored.Length > 0)) {
            decimal storedActual = assignment.ActualCost ?? stored.Sum(v => ProjectXmlValue.ParseMoney(v.Value!));
            if (storedActual != computedActual || stored.Length > 0) {
                charges.RemoveAll(c => c.IsActual);
                if (stored.Length > 0) {
                    foreach (var value in stored) { RequireInterval(value); charges.Add(new ProjectCostInterval(value.Start!.Value, value.Finish!.Value, ProjectXmlValue.ParseMoney(value.Value!), true)); }
                    if (Math.Abs(charges.Where(c => c.IsActual).Sum(c => c.Cost) - storedActual) > .01m)
                        throw new InvalidDataException("Timephased actual costs differ from the stored actual cost.");
                } else if (storedActual != 0) charges.Add(new ProjectCostInterval(start, start, storedActual, true));
                total += storedActual - computedActual; computedActual = storedActual;
            }
        }
        return (total, computedActual);
    }
}
