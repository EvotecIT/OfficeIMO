namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    private sealed class BaselineUpdate {
        internal ProjectCollection<ProjectBaseline> Baselines = null!;
        internal ProjectCollection<ProjectTimephasedValue> Timephased = null!;
        internal int Uid, Number, WorkType, CostType;
        internal DateTime? Start, Finish;
        internal ProjectDuration? Duration;
        internal ProjectWork Work;
        internal decimal? Cost, FixedCost;
        internal readonly List<(int Type, DateTime Start, DateTime Finish, string Value)> Values = new();
        internal void Apply() {
            foreach (var previous in Baselines.Where(b => (b.Number ?? 0) == Number).ToArray()) Baselines.Remove(previous);
            foreach (var value in Timephased.Where(v => v.Type == WorkType || v.Type == CostType).ToArray()) Timephased.Remove(value);
            var baseline = Baselines.Add(); baseline.Number = Number; baseline.Start = Start; baseline.Finish = Finish;
            baseline.Duration = Duration; baseline.Work = Work; baseline.Cost = Cost; baseline.FixedCost = FixedCost;
            foreach (var value in Values) {
                var item = Timephased.Add(); item.Type = value.Type; item.Uid = Uid; item.Start = value.Start;
                item.Finish = value.Finish; item.Unit = 1; item.Value = value.Value;
            }
        }
    }
    /// <summary>Captures task, assignment and resource baseline totals and curves from a current calculated schedule. Existing baselines require explicit overwrite.</summary>
    public void CaptureBaseline(ProjectScheduleResult schedule, int baselineNumber = 0, bool overwrite = false, int maxIntervals = 1_000_000, CancellationToken cancellationToken = default) {
        EnsureMutable();
        if (schedule == null) throw new ArgumentNullException(nameof(schedule));
        if (baselineNumber < 0 || baselineNumber > 10) throw new ArgumentOutOfRangeException(nameof(baselineNumber));
        if (maxIntervals < 1) throw new ArgumentOutOfRangeException(nameof(maxIntervals));
        if (_batchDepth != 0 || schedule.Document != this || schedule.ModelRevision != Revision)
            throw new InvalidOperationException("Baseline capture requires a schedule for the current document revision.");
        if (!schedule.CalculatedAssignments) throw new ArgumentException("Baseline capture requires CalculateAssignments.", nameof(schedule));
        schedule.Report.ThrowIfErrors(); cancellationToken.ThrowIfCancellationRequested();
        foreach (var source in schedule.ExternalSources) source.ValidateCurrent();
        var updates = new List<BaselineUpdate>(); long intervals = 0;
        void AddWork(BaselineUpdate update, ProjectAssignmentInterval interval) {
            if (++intervals > maxIntervals) throw new InvalidOperationException("Baseline capture exceeds MaxIntervals.");
            update.Values.Add((update.WorkType, interval.Start, interval.Finish, ProjectXmlValue.Work(interval.Work)!));
        }
        void AddCost(BaselineUpdate update, ProjectCostInterval interval) {
            if (++intervals > maxIntervals) throw new InvalidOperationException("Baseline capture exceeds MaxIntervals.");
            update.Values.Add((update.CostType, interval.Start, interval.Finish, ProjectXmlValue.Money(interval.Cost)!));
        }
        BaselineUpdate Prepare(ProjectEntity entity, ProjectCollection<ProjectBaseline> baselines, ProjectCollection<ProjectTimephasedValue> phased, (int Work, int Cost) types) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!overwrite && (baselines.Any(b => (b.Number ?? 0) == baselineNumber) || phased.Any(v => v.Type == types.Work || v.Type == types.Cost)))
                throw new InvalidOperationException("The requested baseline already exists; set overwrite explicitly to replace it.");
            var update = new BaselineUpdate { Baselines = baselines, Timephased = phased, Uid = entity.Uid, Number = baselineNumber, WorkType = types.Work, CostType = types.Cost };
            updates.Add(update); return update;
        }
        var assignmentPlans = schedule.Assignments.ToDictionary(a => a.AssignmentUid);
        var assignmentUpdates = new Dictionary<int, BaselineUpdate>();
        foreach (var plan in schedule.Assignments) {
            var assignment = Assignments.GetByUid(plan.AssignmentUid);
            var update = Prepare(assignment, assignment.Baselines, assignment.TimephasedData, ProjectBaselineTypes.Assignment(baselineNumber));
            update.Start = plan.Start; update.Finish = plan.Finish; update.Work = plan.Work; update.Cost = plan.Cost;
            foreach (var interval in plan.Intervals) AddWork(update, interval);
            if (plan.Cost.HasValue) foreach (var interval in plan.Costs) AddCost(update, interval);
            assignmentUpdates.Add(plan.AssignmentUid, update);
        }
        var taskUpdates = new Dictionary<int, BaselineUpdate>();
        foreach (var plan in schedule.Tasks.Reverse()) {
            var task = Tasks.GetByUid(plan.TaskUid);
            var update = Prepare(task, task.Baselines, task.TimephasedData, ProjectBaselineTypes.Task(baselineNumber));
            update.Start = plan.Start; update.Finish = plan.Finish; update.Duration = plan.Duration;
            update.Work = plan.Calculation!.Work; update.Cost = plan.Calculation.Cost; update.FixedCost = task.FixedCost;
            IEnumerable<BaselineUpdate> children;
            if (task.IsSummary) children = (task.Uid == 0 ? Tasks.Where(t => t.Uid != 0) : task.Children).Where(t => taskUpdates.ContainsKey(t.Uid)).Select(t => taskUpdates[t.Uid]);
            else children = Assignments.Where(a => a.Task == task && assignmentUpdates.ContainsKey(a.Uid)).Select(a => assignmentUpdates[a.Uid]);
            foreach (var child in children) foreach (var value in child.Values) {
                if (value.Type == child.WorkType && task.IsSummary == false && Assignments.GetByUid(child.Uid).Resource?.Type != ProjectResourceType.Work) continue;
                if (++intervals > maxIntervals) throw new InvalidOperationException("Baseline capture exceeds MaxIntervals.");
                update.Values.Add((value.Type == child.WorkType ? update.WorkType : update.CostType, value.Start, value.Finish, value.Value));
            }
            if (task.FixedCost is decimal fixedCost && fixedCost != 0) {
                DateTime start = plan.Start, finish = plan.Finish;
                if (task.FixedCostAccrual == ProjectCostAccrual.Start) finish = start;
                else if (task.FixedCostAccrual == ProjectCostAccrual.End) start = finish;
                var active = task.IsSummary || start == finish ? new List<ProjectWorkingRange>() :
                    ProjectCalendarMath.Merge(schedule.Assignments.Where(a => a.TaskUid == task.Uid && Resources.GetByUid(a.ResourceUid).Type == ProjectResourceType.Work)
                        .SelectMany(a => a.Intervals).Where(i => i.Work.Minutes > i.OvertimeWork.Minutes)
                        .Select(i => new ProjectWorkingRange(i.Start, i.Finish)).ToList());
                decimal activeMinutes = active.Sum(r => (r.Finish.Ticks - r.Start.Ticks) / (decimal)TimeSpan.TicksPerMinute);
                decimal durationMinutes = plan.Duration.Value * ProjectXmlValue.MinutesPerUnit(plan.Duration.Unit, plan.Duration.IsElapsed, this);
                if (active.Count == 0 || Math.Abs(activeMinutes - durationMinutes) > .001m)
                    AddCost(update, new ProjectCostInterval(start, finish, fixedCost, false));
                else {
                    decimal span = activeMinutes, assigned = 0m;
                    for (int index = 0; index < active.Count; index++) {
                        var range = active[index];
                        decimal cost = index == active.Count - 1 ? fixedCost - assigned :
                            fixedCost * ((range.Finish.Ticks - range.Start.Ticks) / (decimal)TimeSpan.TicksPerMinute) / span;
                        AddCost(update, new ProjectCostInterval(range.Start, range.Finish, cost, false)); assigned += cost;
                    }
                }
            }
            taskUpdates.Add(task.Uid, update);
        }
        foreach (var group in Assignments.Where(a => a.Resource != null).GroupBy(a => a.Resource!)) {
            if (!group.Any(a => assignmentPlans.ContainsKey(a.Uid))) continue;
            if (group.Any(a => !assignmentPlans.ContainsKey(a.Uid))) throw new InvalidOperationException("Resource baseline capture requires all assignments of each captured resource in the schedule.");
            var plans = group.Select(a => assignmentPlans[a.Uid]).ToArray(); var resource = group.Key;
            var update = Prepare(resource, resource.Baselines, resource.TimephasedData, ProjectBaselineTypes.Resource(baselineNumber));
            update.Work = new ProjectWork(plans.Sum(a => a.Work.Minutes)); update.Cost = plans.All(a => a.Cost.HasValue) ? plans.Sum(a => a.Cost!.Value) : (decimal?)null;
            foreach (var plan in plans) {
                foreach (var interval in plan.Intervals) AddWork(update, interval);
                if (plan.Cost.HasValue) foreach (var interval in plan.Costs) AddCost(update, interval);
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (schedule.ModelRevision != Revision) throw new InvalidOperationException("The document changed during baseline capture.");
        foreach (var source in schedule.ExternalSources) source.ValidateCurrent();
        using (BeginUpdate()) foreach (var update in updates) update.Apply();
    }
}
