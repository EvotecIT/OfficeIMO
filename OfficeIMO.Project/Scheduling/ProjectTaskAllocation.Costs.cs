namespace OfficeIMO.Project;

internal sealed partial class ProjectTaskAllocation {
    private (decimal? Cost, decimal? Actual) CalculateCosts(Entry entry, List<ProjectAssignmentInterval> intervals, DateTime start, DateTime finish, List<ProjectCostInterval> charges) {
        var assignment = entry.Assignment; var resource = assignment.Resource!;
        if (resource.Type == ProjectResourceType.Cost) {
            var recorded = ReadActualCostCurves(assignment);
            if (!assignment.ActualCost.HasValue && recorded.Length == 0 && !(assignment.Cost.HasValue && assignment.RemainingCost.HasValue)
                && (assignment.ActualStart.HasValue || assignment.ActualFinish.HasValue || assignment.Stop.HasValue))
                throw new InvalidDataException("Started cost assignments require actual cost, actual-cost intervals, or both total and remaining cost.");
            decimal actual = assignment.ActualCost ?? (recorded.Length > 0 ? recorded.Sum(c => c.Cost)
                : assignment.Cost.HasValue && assignment.RemainingCost.HasValue ? assignment.Cost.Value - assignment.RemainingCost.Value : 0m);
            decimal? value = assignment.Cost ?? (assignment.RemainingCost.HasValue ? actual + assignment.RemainingCost.Value : (decimal?)null);
            if (value.HasValue && assignment.RemainingCost.HasValue && Math.Abs(value.Value - actual - assignment.RemainingCost.Value) > .01m)
                throw new InvalidDataException("Cost-resource total cost must equal actual plus remaining cost.");
            if (assignment.ActualFinish.HasValue && value.HasValue && Math.Abs(value.Value - actual) > .01m)
                throw new InvalidDataException("A completed cost assignment cannot have remaining cost.");
            if (recorded.Length > 0) {
                if (Math.Abs(recorded.Sum(c => c.Cost) - actual) > .01m)
                    throw new InvalidDataException("Timephased actual costs differ from the stored actual cost.");
                ValidateActualCostDates(assignment, recorded, start, finish);
                charges.AddRange(recorded);
            } else if (value.HasValue || actual != 0) {
                var actualDate = assignment.ActualStart ?? assignment.ActualFinish ?? start;
                charges.Add(new ProjectCostInterval(actualDate, actualDate, actual, true));
            }
            if (value.HasValue) charges.Add(new ProjectCostInterval(finish, finish, value.Value - actual, false));
            return (value, actual);
        }
        var stored = _options.RecalculateActualCosts ? Array.Empty<ProjectCostInterval>() : ReadRetainedActualCosts(assignment, start, finish);
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
            // Failed recalculation retains source actuals, so those records must still be eligible for application.
            if (_options.RecalculateActualCosts) ReadRetainedActualCosts(assignment, start, finish);
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
        if (!_options.RecalculateActualCosts && (assignment.ActualCost.HasValue || stored.Length > 0)) {
            decimal storedActual = assignment.ActualCost ?? stored.Sum(v => v.Cost);
            if (storedActual != computedActual || stored.Length > 0) {
                charges.RemoveAll(c => c.IsActual);
                if (stored.Length > 0) {
                    charges.AddRange(stored);
                    if (Math.Abs(charges.Where(c => c.IsActual).Sum(c => c.Cost) - storedActual) > .01m)
                        throw new InvalidDataException("Timephased actual costs differ from the stored actual cost.");
                } else if (storedActual != 0) charges.Add(new ProjectCostInterval(start, start, storedActual, true));
                total += storedActual - computedActual; computedActual = storedActual;
            }
        }
        return (total, computedActual);
    }
    private ProjectCostInterval[] ReadActualCostCurves(ProjectAssignment assignment) {
        var values = new List<ProjectCostInterval>();
        foreach (var item in assignment.TimephasedData.Where(v => v.Type == 6)) {
            _token.ThrowIfCancellationRequested(); RequireInterval(item);
            values.Add(new ProjectCostInterval(item.Start!.Value, item.Finish!.Value, ProjectXmlValue.ParseMoney(item.Value!), true));
            CheckCount(values.Count);
        }
        return values.ToArray();
    }

    private void ValidateActualCostDates(ProjectAssignment assignment, IEnumerable<ProjectCostInterval> intervals, DateTime start, DateTime finish) {
        foreach (var interval in intervals) {
            _token.ThrowIfCancellationRequested();
            if (interval.Start < start || interval.Finish > finish
                || (assignment.ActualStart.HasValue && interval.Start < assignment.ActualStart)
                || (assignment.ActualFinish.HasValue && interval.Finish > assignment.ActualFinish)
                || (assignment.Stop.HasValue && interval.Finish > assignment.Stop))
                throw new InvalidDataException("Actual-cost intervals must fit the task span and recorded assignment actual dates.");
        }
    }

    private ProjectCostInterval[] ReadRetainedActualCosts(ProjectAssignment assignment, DateTime start, DateTime finish) {
        var stored = ReadActualCostCurves(assignment);
        ValidateActualCostDates(assignment, stored, start, finish);
        if (stored.Length > 0 && assignment.ActualCost.HasValue && Math.Abs(stored.Sum(v => v.Cost) - assignment.ActualCost.Value) > .01m)
            throw new InvalidDataException("Timephased actual costs differ from the stored actual cost.");
        return stored;
    }
}
