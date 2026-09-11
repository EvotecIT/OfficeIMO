namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    private sealed class AssignmentUpdate {
        internal ProjectAssignment Target = null!;
        internal ProjectAssignmentSchedule Plan = null!;
        internal readonly HashSet<int> ReplacedTypes = new() { 1 };
        internal readonly List<(int Type, DateTime Start, DateTime Finish, string Value)> Values = new();
        internal void Apply() {
            Target.Start = Plan.Start; Target.Finish = Plan.Finish; Target.Units = Plan.Units;
            Target.Work = Plan.Work; Target.ActualWork = Plan.ActualWork; Target.RemainingWork = Plan.RemainingWork;
            Target.OvertimeWork = Plan.OvertimeWork; Target.ActualOvertimeWork = Plan.ActualOvertimeWork;
            Target.Cost = Plan.Cost; Target.ActualCost = Plan.ActualCost; Target.RemainingCost = Plan.RemainingCost;
            Target.PercentWorkComplete = Plan.Work.Minutes == 0 ? 0 : (int)decimal.Round(Plan.ActualWork.Minutes / Plan.Work.Minutes * 100m, 0, MidpointRounding.AwayFromZero);
            foreach (var value in Target.TimephasedData.Where(v => v.Type.HasValue && ReplacedTypes.Contains(v.Type.Value)).ToArray()) Target.TimephasedData.Remove(value);
            foreach (var value in Values) {
                var item = Target.TimephasedData.Add(); item.Type = value.Type; item.Uid = Target.Uid;
                item.Start = value.Start; item.Finish = value.Finish; item.Unit = 1; item.Value = value.Value;
            }
        }
    }
    private AssignmentUpdate[] PrepareAssignmentUpdates(ProjectScheduleResult result) {
        var updates = new List<AssignmentUpdate>();
        foreach (var plan in result.Assignments) {
            var target = Assignments.SingleOrDefault(a => a.Uid == plan.AssignmentUid)
                ?? throw new InvalidOperationException("A scheduled assignment is no longer attached.");
            if (target.Task?.Uid != plan.TaskUid || target.Resource?.Uid != plan.ResourceUid)
                throw new InvalidOperationException("A scheduled assignment's relationships changed.");
            var update = new AssignmentUpdate { Target = target, Plan = plan };
            // Scalar-authored progress needs a compact actual-work curve for independent importers.
            // Preserve existing source actual records and their identities when they already exist.
            if (!target.TimephasedData.Any(v => v.Type == 2)) {
                foreach (var interval in plan.Intervals.Where(i => i.IsActual)) {
                    update.Values.Add((2, interval.Start, interval.Finish, ProjectXmlValue.Work(interval.Work)!));
                    if (interval.OvertimeWork.Minutes > 0 && !target.TimephasedData.Any(v => v.Type == 3))
                        update.Values.Add((3, interval.Start, interval.Finish, ProjectXmlValue.Work(interval.OvertimeWork)!));
                }
            }
            var intervals = plan.Intervals.Where(i => !i.IsActual).OrderBy(i => i.Start).ToArray();
            DateTime? lastFinish = plan.RemainingAnchor;
            foreach (var interval in intervals) {
                // Explicit zero records retain interruptions when these curves are scheduled again.
                if (lastFinish.HasValue && interval.Start > lastFinish.Value)
                    update.Values.Add((1, lastFinish.Value, interval.Start, "PT0H0M0S"));
                update.Values.Add((1, interval.Start, interval.Finish,
                    ProjectXmlValue.Work(new ProjectWork(interval.Work.Minutes - interval.OvertimeWork.Minutes))!));
                lastFinish = interval.Finish;
            }
            if (plan.ActualCost != target.ActualCost) update.ReplacedTypes.Add(6);
            if (update.ReplacedTypes.Contains(6))
                foreach (var charge in plan.Costs.Where(c => c.IsActual))
                    update.Values.Add((6, charge.Start, charge.Finish, ProjectXmlValue.Money(charge.Cost)!));
            updates.Add(update);
        }
        return updates.ToArray();
    }
    private sealed class ResourceUpdate {
        internal ProjectResource Target = null!;
        internal decimal? Work, ActualWork, RemainingWork, Cost, ActualCost;
        internal void Apply() {
            Target.Work = Work.HasValue ? new ProjectWork(Work.Value) : (ProjectWork?)null;
            Target.ActualWork = ActualWork.HasValue ? new ProjectWork(ActualWork.Value) : (ProjectWork?)null;
            Target.RemainingWork = RemainingWork.HasValue ? new ProjectWork(RemainingWork.Value) : (ProjectWork?)null;
            Target.Cost = Cost; Target.ActualCost = ActualCost;
        }
    }
    private ResourceUpdate[] PrepareResourceUpdates(ProjectScheduleResult result) {
        if (!result.CalculatedAssignments) return Array.Empty<ResourceUpdate>();
        var plans = result.Assignments.ToDictionary(a => a.AssignmentUid);
        var updates = new List<ResourceUpdate>();
        decimal? Sum(IEnumerable<decimal?> values) { decimal total = 0; foreach (var value in values) { if (!value.HasValue) return null; total = checked(total + value.Value); } return total; }
        foreach (var group in Assignments.Where(a => a.Resource != null).GroupBy(a => a.Resource!)) {
            if (!group.Any(a => plans.ContainsKey(a.Uid))) continue;
            var entries = group.Select(a => (Source: a, Plan: plans.TryGetValue(a.Uid, out var plan) ? plan : null)).ToArray();
            updates.Add(new ResourceUpdate { Target = group.Key,
                Work = Sum(entries.Select(e => e.Plan == null ? e.Source.Work?.Minutes : e.Plan.Work.Minutes)),
                ActualWork = Sum(entries.Select(e => e.Plan == null ? e.Source.ActualWork?.Minutes : e.Plan.ActualWork.Minutes)),
                RemainingWork = Sum(entries.Select(e => e.Plan == null ? e.Source.RemainingWork?.Minutes : e.Plan.RemainingWork.Minutes)),
                Cost = Sum(entries.Select(e => e.Plan == null ? e.Source.Cost : e.Plan.Cost)),
                ActualCost = Sum(entries.Select(e => e.Plan == null ? e.Source.ActualCost : e.Plan.ActualCost)) });
        }
        return updates.ToArray();
    }
}
