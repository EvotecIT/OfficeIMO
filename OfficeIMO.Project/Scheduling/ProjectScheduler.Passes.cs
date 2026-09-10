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
                var predecessor = _nodes[link.Predecessor!];
                DateTime bound = Lag(link, FromFinish(link) ? predecessor.EarlyFinish : predecessor.EarlyStart, false);
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
            node.EarlyStart = Snap(node, start, true); node.EarlyFinish = Add(node, node.EarlyStart, node.Minutes);
            if (task.IsManual == true) { node.EarlyStart = task.Start!.Value; node.EarlyFinish = task.Finish!.Value; }
        }
    }
    private void Backward(DateTime horizon) {
        ProjectCalendarMath.Local(horizon);
        foreach (var node in _order.AsEnumerable().Reverse()) {
            _token.ThrowIfCancellationRequested();
            var task = node.Task;
            DateTime finish = Snap(node, horizon, false);
            if (task.Deadline.HasValue) finish = Min(finish, task.Deadline.Value);
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
                var predecessor = _nodes[link.Predecessor!];
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
                node.LateStart, node.LateFinish, total, free, total <= _options.CriticalSlackMinutes, false, ResultDuration(node.Task, node.Minutes, node.Elapsed)));
        }
        var all = _document.AllTasks.ToArray();
        foreach (var task in all.AsEnumerable().Reverse()) {
            _token.ThrowIfCancellationRequested();
            if (!task.IsSummary || task.IsActive == false) continue;
            var children = (task.Uid == 0 ? _document.Tasks.Where(t => t.Uid != 0) : task.Children).Where(results.ContainsKey).Select(t => results[t]).ToArray();
            if (children.Length == 0) continue;
            results.Add(task, new ProjectTaskSchedule(task.Uid, children.Min(c => c.Start), children.Max(c => c.Finish),
                children.Min(c => c.EarlyStart), children.Max(c => c.EarlyFinish), children.Min(c => c.LateStart), children.Max(c => c.LateFinish),
                children.Min(c => c.TotalSlackMinutes), children.Min(c => c.FreeSlackMinutes), children.Any(c => c.IsCritical), true,
                ResultDuration(task, CalendarMath(new[] { task.Calendar ?? _document.Calendar ?? throw new InvalidOperationException("Summary rollup requires a project calendar.") })
                    .Between(children.Min(c => c.Start), children.Max(c => c.Finish)), false)));
        }
        return all.Where(results.ContainsKey).Select(t => results[t]).ToArray();
    }
    private static decimal MinutesBetween(Node node, DateTime start, DateTime finish) => node.Elapsed
        ? (finish.Ticks - start.Ticks) / (decimal)TimeSpan.TicksPerMinute : node.Calendar.Between(start, finish);
    private ProjectDuration ResultDuration(ProjectTask task, decimal minutes, bool elapsed) {
        var unit = task.Duration?.Unit ?? ProjectDurationUnit.Day;
        return new ProjectDuration(minutes / ProjectXmlValue.MinutesPerUnit(unit, elapsed, _document), unit, elapsed, task.Duration?.IsEstimated ?? false);
    }
}
