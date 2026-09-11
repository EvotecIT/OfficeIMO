namespace OfficeIMO.Project;

internal sealed partial class ProjectScheduler {
    private void Forward(DateTime origin) {
        ProjectCalendarMath.Local(origin);
        foreach (var node in _order) {
            _token.ThrowIfCancellationRequested();
            var task = node.Task;
            DateTime start = Snap(node, origin, true);
            bool hasDependencyBound = false;
            foreach (var link in node.In) {
                var predecessor = PredecessorBounds(link, true);
                DateTime bound = Lag(link, FromFinish(link) ? predecessor.Finish : predecessor.Start, false);
                DateTime candidate = ToFinish(link) ? Add(node, bound, -node.Minutes) : bound;
                start = hasDependencyBound ? Max(start, candidate) : candidate;
                hasDependencyBound = true;
            }
            if (task.ConstraintDate.HasValue) {
                var date = task.ConstraintDate.Value;
                switch (task.ConstraintType) {
                    case ProjectConstraintType.StartNoEarlierThan: start = Max(start, date); break;
                    case ProjectConstraintType.FinishNoEarlierThan: start = Max(start, Add(node, date, -node.Minutes)); break;
                    case ProjectConstraintType.MustStartOn: start = date; break;
                    case ProjectConstraintType.MustFinishOn: start = Add(node, date, -node.Minutes); break;
                }
            }
            node.BeforeLevelingAnchor = start;
            if (task.LevelingDelay is ProjectDuration leveling && leveling.Value > 0) {
                decimal delay = leveling.Value * ProjectXmlValue.MinutesPerUnit(leveling.Unit, leveling.IsElapsed, _document);
                start = leveling.IsElapsed ? start.Add(ProjectXmlValue.MinutesToSpan(delay)) : node.Calendar.Add(start, delay);
            }
            if (_notBefore != null && _notBefore.TryGetValue(task.Uid, out var minimum)) start = Max(start, minimum);
            node.EarlyStart = Snap(node, start, true); node.EarlyFinish = Add(node, node.EarlyStart, node.Minutes);
            node.EarlyAnchor = node.EarlyStart;
            if (node.Allocation != null) {
                var allocation = node.Allocation.Build(node.EarlyStart, true);
                node.EarlyStart = allocation.Start; node.EarlyFinish = allocation.Finish; node.Minutes = allocation.Duration;
            }
            if (task.IsManual == true) { node.EarlyStart = task.Start!.Value; node.EarlyFinish = task.Finish!.Value; }
        }
    }
    private void Backward(DateTime horizon) {
        ProjectCalendarMath.Local(horizon);
        foreach (var node in _order.AsEnumerable().Reverse()) {
            _token.ThrowIfCancellationRequested();
            var task = node.Task;
            if (node.Allocation?.HasActuals == true) { node.LateStart = node.EarlyStart; node.LateFinish = node.EarlyFinish; node.LateAnchor = node.EarlyAnchor; continue; }
            DateTime finish = Snap(node, horizon, false);
            if (task.Deadline.HasValue) finish = Min(finish, task.Deadline.Value);
            // Percentage lags need assignment-derived effort duration before predecessor inversion.
            if (node.Allocation != null && node.Out.Any(link => link.LagPercent.HasValue))
                node.Minutes = node.Allocation.Build(finish, false).Duration;
            foreach (var link in node.Out) {
                var successor = _nodes[link.Successor];
                DateTime bound = Lag(link, ToFinish(link) ? successor.LateFinish : successor.LateStart, true);
                finish = Min(finish, FromFinish(link) ? bound : Add(node, bound, node.Minutes));
            }
            if (task.ConstraintDate.HasValue) {
                var date = task.ConstraintDate.Value;
                switch (task.ConstraintType) {
                    case ProjectConstraintType.StartNoLaterThan: finish = Min(finish, Add(node, date, node.Minutes)); break;
                    case ProjectConstraintType.FinishNoLaterThan: finish = Min(finish, date); break;
                    case ProjectConstraintType.MustStartOn: finish = Add(node, date, node.Minutes); break;
                    case ProjectConstraintType.MustFinishOn: finish = date; break;
                }
            }
            node.LateFinish = Snap(node, finish, false); node.LateStart = Add(node, node.LateFinish, -node.Minutes);
            node.LateAnchor = node.LateStart;
            if (node.Allocation != null) {
                var allocation = node.Allocation.Build(node.LateFinish, false);
                node.LateStart = allocation.Start; node.LateFinish = allocation.Finish; node.Minutes = allocation.Duration;
                node.LateAnchor = allocation.Start;
            }
            if (task.ConstraintType == ProjectConstraintType.StartNoLaterThan && task.ConstraintDate.HasValue)
                node.LateStart = Min(node.LateStart, task.ConstraintDate.Value);
            if (task.IsManual == true) { node.LateStart = task.Start!.Value; node.LateFinish = task.Finish!.Value; }
        }
    }
    private void CheckFinalBounds() {
        foreach (var node in _order) {
            _token.ThrowIfCancellationRequested();
            var task = node.Task;
            if (task.ConstraintDate is DateTime date) {
                bool conflict = task.ConstraintType switch {
                    ProjectConstraintType.MustStartOn => node.Start != date,
                    ProjectConstraintType.MustFinishOn => node.Finish != date,
                    ProjectConstraintType.StartNoEarlierThan => node.Start < date,
                    ProjectConstraintType.StartNoLaterThan => node.Start > date,
                    ProjectConstraintType.FinishNoEarlierThan => node.Finish < date,
                    ProjectConstraintType.FinishNoLaterThan => node.Finish > date,
                    _ => false
                };
                if (conflict) Error("PROJECT_CONSTRAINT_CONFLICT", "The calculated dates cannot satisfy the task's date constraint.", task);
            }
            foreach (var link in node.In) {
                var predecessor = PredecessorBounds(link, false);
                DateTime bound = Lag(link, FromFinish(link) ? predecessor.Finish : predecessor.Start, false);
                DateTime target = ToFinish(link) ? node.Finish : node.Start;
                // A finish at one shift boundary and the following shift's start represent the same working-time boundary.
                if (target < bound && (node.Elapsed || node.Calendar.Between(target, bound) > 0))
                    Error("PROJECT_DEPENDENCY_CONFLICT", "A manual date or hard/late constraint conflicts with a predecessor bound.", task);
            }
            if (task.Deadline.HasValue && node.Finish > task.Deadline.Value)
                _diagnostics.Add(new ProjectDiagnostic("PROJECT_DEADLINE_MISSED", ProjectDiagnosticSeverity.Warning,
                    "The proposed finish exceeds the deadline; total float includes this target.", "/Task[UID=" + task.Uid + "]"));
        }
    }
    private IEnumerable<ProjectTaskSchedule> BuildResults(DateTime horizon) {
        var results = new Dictionary<ProjectTask, ProjectTaskSchedule>();
        foreach (var node in _order) {
            _token.ThrowIfCancellationRequested();
            decimal total = node.Task.ConstraintType == ProjectConstraintType.AsLateAsPossible ? 0m : MinutesBetween(node, node.EarlyStart, node.LateStart);
            if (node.Allocation != null) node.Plan = node.Allocation.Build(node.PlanAnchor, true, true);
            if (node.Plan != null && (node.Plan.Start != node.Start || node.Plan.Finish != node.Finish))
                Error("PROJECT_ASSIGNMENT_BOUND_CONFLICT", "Final assignment intervals differ from the task boundaries established by scheduling.", node.Task);
            decimal free = MinutesBetween(node, node.Finish, horizon);
            foreach (var link in node.Out) {
                var successor = _nodes[link.Successor];
                DateTime bound = Lag(link, ToFinish(link) ? successor.Finish : successor.Start, true);
                DateTime anchor = FromFinish(link) ? node.Finish : node.Start;
                free = Math.Min(free, MinutesBetween(node, anchor, bound));
            }
            if (node.Task.ConstraintType is ProjectConstraintType.MustStartOn or ProjectConstraintType.MustFinishOn) free = 0;
            else if (node.Task.ConstraintType is ProjectConstraintType.StartNoLaterThan or ProjectConstraintType.FinishNoLaterThan) free = Math.Min(free, total);
            results.Add(node.Task, new ProjectTaskSchedule(node.Task.Uid, node.Start, node.Finish, node.EarlyStart, node.EarlyFinish,
                node.LateStart, node.LateFinish, total, free, total <= _options.CriticalSlackMinutes, false,
                ResultDuration(node.Task, node.Plan?.Duration ?? node.Minutes, node.Elapsed), node.Plan == null ? null : node.Allocation!.Totals(node.Plan),
                node.BeforeLevelingAnchor, node.EarlyAnchor));
        }
        var all = _document.AllTasks.ToArray();
        // Source order does not constrain the project summary; calculate it after ordinary summaries.
        foreach (var task in all.AsEnumerable().Reverse().Where(t => t.Uid != 0).Concat(all.Where(t => t.Uid == 0))) {
            _token.ThrowIfCancellationRequested();
            if (!task.IsSummary || task.IsActive == false) continue;
            var children = (task.Uid == 0 ? _document.Tasks.Where(t => t.Uid != 0) : task.Children).Where(results.ContainsKey).Select(t => results[t]).ToArray();
            if (children.Length == 0) continue;
            decimal summaryMinutes = CalendarMath(new[] { task.Calendar ?? _document.Calendar ?? throw new InvalidOperationException("Summary rollup requires a project calendar.") })
                .Between(children.Min(c => c.Start), children.Max(c => c.Finish));
            results.Add(task, new ProjectTaskSchedule(task.Uid, children.Min(c => c.Start), children.Max(c => c.Finish),
                children.Min(c => c.EarlyStart), children.Max(c => c.EarlyFinish), children.Min(c => c.LateStart), children.Max(c => c.LateFinish),
                children.Min(c => c.TotalSlackMinutes), children.Min(c => c.FreeSlackMinutes), children.Any(c => c.IsCritical), true,
                ResultDuration(task, summaryMinutes, false), _options.CalculateAssignments ? SummaryTotals(task, children, summaryMinutes) : null));
        }
        return all.Where(results.ContainsKey).Select(t => results[t]).ToArray();
    }
    private static ProjectTaskWorkSchedule SummaryTotals(ProjectTask task, ProjectTaskSchedule[] children, decimal duration) {
        var totals = children.Select(c => c.Calculation!).ToArray();
        decimal work = totals.Sum(c => c.Work.Minutes), actual = totals.Sum(c => c.ActualWork.Minutes);
        decimal childDuration = totals.Sum(c => c.ActualDuration.Value + c.RemainingDuration.Value);
        bool completed = totals.All(c => c.PercentComplete == 100);
        decimal fraction = childDuration == 0 ? completed ? 1m : 0m : totals.Sum(c => c.ActualDuration.Value) / childDuration;
        decimal fixedCost = task.FixedCost ?? 0m;
        decimal actualFixed = (task.FixedCostAccrual ?? ProjectCostAccrual.Prorated) switch {
            ProjectCostAccrual.Start => fraction > 0 ? fixedCost : 0m,
            ProjectCostAccrual.End => fraction == 1m ? fixedCost : 0m,
            _ => fixedCost * fraction
        };
        return new ProjectTaskWorkSchedule(new ProjectWork(work), new ProjectWork(actual), new ProjectWork(work - actual), duration * fraction, duration * (1m - fraction),
            totals.All(c => c.Cost.HasValue) ? totals.Sum(c => c.Cost!.Value) + fixedCost : (decimal?)null,
            totals.All(c => c.ActualCost.HasValue) ? totals.Sum(c => c.ActualCost!.Value) + actualFixed : (decimal?)null, task.PhysicalPercentComplete, completed: completed);
    }
    private static decimal MinutesBetween(Node node, DateTime start, DateTime finish) => node.Elapsed
        ? (finish.Ticks - start.Ticks) / (decimal)TimeSpan.TicksPerMinute : node.Calendar.Between(start, finish);
    private ProjectDuration ResultDuration(ProjectTask task, decimal minutes, bool elapsed) {
        var unit = task.Duration?.Unit ?? ProjectDurationUnit.Day;
        return new ProjectDuration(minutes / ProjectXmlValue.MinutesPerUnit(unit, elapsed, _document), unit, elapsed, task.Duration?.IsEstimated ?? false);
    }
}
