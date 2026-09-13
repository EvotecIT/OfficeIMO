namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    internal void CheckCalculatedCostReplacement(IReadOnlyList<ProjectTaskSchedule> tasks, IReadOnlyList<ProjectAssignmentSchedule> assignments,
        Action<ProjectDiagnostic> add, CancellationToken token) {
        void Check(decimal? stored, bool calculatedKnown, string location) {
            if (stored.HasValue && !calculatedKnown)
                add(new ProjectDiagnostic("PROJECT_COST_RECALCULATION_INCOMPLETE", ProjectDiagnosticSeverity.Error,
                    "The calculated cost is unknown and cannot replace an entered amount. Supply complete rate and cost inputs before applying this schedule.", location));
        }
        foreach (var plan in tasks) {
            token.ThrowIfCancellationRequested();
            if (plan.Calculation is not ProjectTaskWorkSchedule cost) continue;
            var task = TaskIndex[plan.TaskUid]; string location = "/Task[UID=" + task.Uid + "]";
            Check(task.Cost, cost.Cost.HasValue, location + "/Cost");
            Check(task.ActualCost, cost.ActualCost.HasValue, location + "/ActualCost");
            Check(task.RemainingCost, cost.RemainingCost.HasValue, location + "/RemainingCost");
        }
        var plans = assignments.ToDictionary(a => a.AssignmentUid);
        foreach (var plan in assignments) {
            token.ThrowIfCancellationRequested();
            var assignment = Assignments.GetByUid(plan.AssignmentUid); string location = "/Assignment[UID=" + assignment.Uid + "]";
            Check(assignment.Cost, plan.Cost.HasValue, location + "/Cost");
            Check(assignment.ActualCost, plan.ActualCost.HasValue, location + "/ActualCost");
            Check(assignment.RemainingCost, plan.RemainingCost.HasValue, location + "/RemainingCost");
        }
        foreach (var group in Assignments.Where(a => a.Resource != null).GroupBy(a => a.Resource!)) {
            token.ThrowIfCancellationRequested();
            if (!group.Any(a => plans.ContainsKey(a.Uid))) continue;
            bool costKnown = true, actualKnown = true, remainingKnown = true;
            foreach (var assignment in group) {
                token.ThrowIfCancellationRequested();
                plans.TryGetValue(assignment.Uid, out var plan);
                costKnown &= (plan == null ? assignment.Cost : plan.Cost).HasValue;
                actualKnown &= (plan == null ? assignment.ActualCost : plan.ActualCost).HasValue;
                remainingKnown &= (plan == null ? assignment.RemainingCost : plan.RemainingCost).HasValue;
            }
            string location = "/Resource[UID=" + group.Key.Uid + "]";
            Check(group.Key.Cost, costKnown, location + "/Cost");
            Check(group.Key.ActualCost, actualKnown, location + "/ActualCost");
            Check(group.Key.RemainingCost, remainingKnown, location + "/RemainingCost");
        }
    }
}
