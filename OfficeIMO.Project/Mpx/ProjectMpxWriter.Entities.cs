namespace OfficeIMO.Project;

internal sealed partial class ProjectMpxWriter {
    private void Resources() {
        var customs = ProjectMpxFields.CustomMappings(false).ToArray();
        var fields = new[] { 40, 49 }.Concat(ProjectMpxFields.Resources.Where(f => f.Id != 40 && f.Id != 10).Select(f => f.Id)).Concat(new[] { 21, 31, 48 }).Concat(customs.Select(f => f.Id)).ToArray();
        Record(new[] { "41" }.Concat(fields.Select(f => ProjectMpxValues.Text(f))).ToArray());
        var declared = _document.Resources.Where(r => r.DisplayId.HasValue).Select(r => r.DisplayId!.Value).ToArray();
        var used = new HashSet<int>(declared); int sequence = 1;
        if (used.Count != declared.Length || declared.Any(row => row < 0 || row > 9999)) throw new InvalidDataException("MPX resource display IDs must be unique and between 0 and 9999.");
        foreach (var resource in _document.Resources) {
            _token.ThrowIfCancellationRequested(); string path = Path(resource, "Resource");
            while (used.Contains(sequence)) sequence++;
            int row = resource.DisplayId ?? sequence;
            if (row > 9999) throw new InvalidDataException("MPX resource display IDs exceed 9999.");
            used.Add(row);
            _resourceRows.Add(resource, row); Handle(path + "/DisplayId"); Handle(path + "/Uid");
            if (resource.Type == ProjectResourceType.Work) Handle(path + "/Type");
            else if (!resource.Type.HasValue) Diagnostic("PROJECT_MPX_RESOURCE_DEFAULT", "MPX normalizes an absent resource type to work.", path + "/Type");
            var values = new Dictionary<int, string> { [40] = ProjectMpxValues.Text(row), [49] = ProjectMpxValues.Text(resource.Uid) };
            foreach (var field in ProjectMpxFields.Resources.Where(f => f.Id != 40)) { values[field.Id] = field.Write(resource); Handle(path + "/" + field.ModelKey); }
            if (resource.Calendar != null) { var calendar = resource.Calendar; values[48] = _calendarNames[_resourceCalendarOwners.ContainsKey(calendar) ? calendar.BaseCalendar! : calendar]; }
            Baseline(resource.Baselines, values, path, false); Custom(resource.CustomFields, values, customs, path);
            Record(new[] { "50" }.Concat(fields.Select(f => values.TryGetValue(f, out var v) ? v : "")).ToArray());
            if (resource.Notes != null) Record("51", ProjectMpxFields.Notes(resource.Notes));
            ResourceCalendar(resource, path);
        }
    }
    private void Tasks() {
        var tasks = _document.AllTasks.ToArray();
        bool preserveRows = true; int previous = -1;
        foreach (var task in tasks) {
            if (!task.DisplayId.HasValue || task.DisplayId <= previous || task.DisplayId > 9999) { preserveRows = false; break; }
            previous = task.DisplayId.Value;
        }
        int next = tasks.Length != 0 && tasks[0].SourceOutlineLevel == 0 ? 0 : 1;
        foreach (var task in tasks) _taskRows.Add(task, preserveRows ? task.DisplayId!.Value : next++);
        var customs = ProjectMpxFields.CustomMappings(true).ToArray();
        var fields = new[] { 90, 98, 3, 120 }.Concat(ProjectMpxFields.Tasks.Where(f => f.Id != 90 && f.Id != 14).Select(f => f.Id)).Concat(new[] { 21, 31, 41, 56, 57, 74 }).Concat(customs.Select(f => f.Id)).ToArray();
        Record(new[] { "61" }.Concat(fields.Select(f => ProjectMpxValues.Text(f))).ToArray());
        var incoming = _document.Dependencies.Select((d, i) => (Dependency: d, Index: i)).GroupBy(p => p.Dependency.Successor).ToDictionary(g => g.Key, g => g.ToArray());
        var assignments = _document.Assignments.GroupBy(a => a.Task!).ToDictionary(g => g.Key, g => g.ToArray());
        int assignmentUid = 0;
        foreach (var task in tasks) {
            _token.ThrowIfCancellationRequested(); string path = Path(task, "Task");
            int level = 1; for (var parent = task.Parent; parent != null; parent = parent.Parent) level++;
            if (task.Parent == null && task.SourceOutlineLevel == 0) level = 0;
            var values = new Dictionary<int, string> {
                [90] = ProjectMpxValues.Text(_taskRows[task]), [98] = ProjectMpxValues.Text(task.Uid), [3] = ProjectMpxValues.Text(level), [120] = ProjectMpxValues.Text(task.IsSummary)
            };
            Handle(path + "/Uid"); Handle(path + "/DisplayId"); Handle(path + "/Parent"); Handle(path + "/Position"); Handle(path + "/IsSummary");
            if (task.DisplayId.HasValue && task.DisplayId != _taskRows[task]) Diagnostic("PROJECT_MPX_TASK_ROWS", "Display rows are renumbered to retain the current outline order; stable task UIDs and relationships are retained.", path + "/DisplayId");
            foreach (var field in ProjectMpxFields.Tasks.Where(f => f.Id != 90)) { values[field.Id] = field.Write(task); Handle(path + "/" + field.ModelKey); }
            if (task.Type == ProjectTaskType.FixedWork) Diagnostic("PROJECT_MPX_TASK_TYPE", "MPX Fixed distinguishes fixed duration from other task types; fixed work becomes fixed units.", path + "/Type");
            if (task.Priority.HasValue) {
                int represented = Math.Max(100, Math.Min(1000, (int)decimal.Round(task.Priority.Value / 100m, 0, MidpointRounding.AwayFromZero) * 100));
                if (represented != task.Priority.Value) Diagnostic("PROJECT_MPX_PRIORITY", "MPX priority classes round this value to " + represented + ".", path + "/Priority");
            }
            Baseline(task.Baselines, values, path, true); Custom(task.CustomFields, values, customs, path);
            if (incoming.TryGetValue(task, out var links)) {
                var expressions = new List<string>();
                foreach (var link in links) {
                    var d = link.Dependency; string key = "/Dependency[" + link.Index + "]";
                    if (d.Predecessor == null || d.CrossProject == true || d.CrossProjectName != null) continue;
                    HandleTree(key);
                    if (d.LagPercentIsElapsed || d.LagPercentIsEstimated)
                        Diagnostic("PROJECT_MPX_PERCENTAGE_LAG_FORMAT", "MPX output does not represent elapsed or estimated percentage lag; the relationship uses working percentage lag.", key + "/LagPercent");
                    string type = d.Type switch { ProjectDependencyType.FinishToFinish => "FF", ProjectDependencyType.StartToStart => "SS", ProjectDependencyType.StartToFinish => "SF", ProjectDependencyType.FinishToStart => "FS", _ => "" };
                    string lag = d.LagPercent.HasValue ? ProjectMpxValues.Text(d.LagPercent.Value) + "%" : ProjectMpxValues.Text(d.Lag);
                    if (lag.Length != 0 && lag[0] != '-') lag = "+" + lag;
                    expressions.Add(ProjectMpxValues.Text(d.Predecessor.Uid) + type + lag);
                }
                values[74] = string.Join(_separator.ToString(), expressions);
            }
            Record(new[] { "70" }.Concat(fields.Select(f => values.TryGetValue(f, out var v) ? v : "")).ToArray());
            if (task.Notes != null) Record("71", ProjectMpxFields.Notes(task.Notes));
            if (!assignments.TryGetValue(task, out var assigned)) continue;
            if (assigned.Length > 100) throw new NotSupportedException("MPX supports at most 100 assignments per task.");
            foreach (var assignment in assigned) Assignment(assignment, ++assignmentUid);
        }
    }
    private void Baseline(ProjectCollection<ProjectBaseline> baselines, Dictionary<int, string> values, string parent, bool task) {
        Handle(parent + "/Baseline/Count");
        for (int i = 0; i < baselines.Count; i++) {
            var baseline = baselines[i]; if (baseline.Number != 0) continue;
            string path = parent + "/Baseline[" + i + "]"; Handle(path + "/Number");
            values[21] = Value(path + "/Work", baseline.Work); values[31] = Value(path + "/Cost", baseline.Cost);
            if (!task) continue;
            values[41] = Value(path + "/Duration", baseline.Duration); values[56] = Value(path + "/Start", baseline.Start); values[57] = Value(path + "/Finish", baseline.Finish);
        }
    }
    private void Custom(ProjectCollection<ProjectCustomFieldValue> customFields, Dictionary<int, string> values, ProjectMpxFields.CustomMapping[] mappings, string parent) {
        Handle(parent + "/Custom/Count"); var used = new HashSet<int>();
        for (int i = 0; i < customFields.Count; i++) {
            var field = customFields[i]; var mapping = mappings.FirstOrDefault(m => m.FieldId == field.FieldId);
            if (mapping == null || field.ValueGuid != null || field.ValueId != null) continue;
            if (!used.Add(mapping.Id)) throw new InvalidDataException("MPX scalar custom fields require one value per entity and field.");
            values[mapping.Id] = mapping.Write(field, _values);
            string path = parent + "/Custom[" + i + "]"; Handle(path + "/FieldId"); Handle(path + "/Value");
            if (mapping.Kind == "Duration") Handle(path + "/DurationFormat");
        }
    }
    private void Assignment(ProjectAssignment item, int uid) {
        _token.ThrowIfCancellationRequested(); string path = Path(item, "Assignment");
        if (item.Resource == null || item.Task == null) throw new InvalidDataException("MPX assignments require resolved task and resource references.");
        Handle(path + "/Uid"); Handle(path + "/Task"); Handle(path + "/Resource"); Handle(path + "/Baseline/Count");
        if (item.Uid != uid) Diagnostic("PROJECT_MPX_ASSIGNMENT_ID", "MPX has no assignment UID field; assignment identities are assigned again in task order.", path + "/Uid");
        ProjectBaseline? baseline = null; string baselinePath = "";
        for (int i = 0; i < item.Baselines.Count; i++) if (item.Baselines[i].Number == 0) { baseline = item.Baselines[i]; baselinePath = path + "/Baseline[" + i + "]"; Handle(baselinePath + "/Number"); }
        Record("75", ProjectMpxValues.Text(_resourceRows[item.Resource]), Value(path + "/Units", item.Units), Value(path + "/Work", item.Work),
            baseline == null ? "" : Value(baselinePath + "/Work", baseline.Work), Value(path + "/ActualWork", item.ActualWork), Value(path + "/OvertimeWork", item.OvertimeWork),
            Value(path + "/Cost", item.Cost), baseline == null ? "" : Value(baselinePath + "/Cost", baseline.Cost), Value(path + "/ActualCost", item.ActualCost),
            Value(path + "/Start", item.Start), Value(path + "/Finish", item.Finish), "", ProjectMpxValues.Text(item.Resource.Uid));
    }
}
