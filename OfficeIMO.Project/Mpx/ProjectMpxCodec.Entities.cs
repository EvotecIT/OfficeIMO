using System.Text.RegularExpressions;

namespace OfficeIMO.Project;

internal static partial class ProjectMpxCodec {
    private sealed partial class Reader {
        private Dictionary<int, string> Values(string[] r, int[]? fields, bool task) {
            if (fields == null) throw new InvalidDataException("An MPX entity appears before its field definition.");
            var values = new Dictionary<int, string>();
            for (int i = 0; i < fields.Length; i++) {
                string text = Get(r, i + 1); if (text.Length == 0) continue;
                int id = fields[i];
                if (ProjectMpxValues.Empty(text)) {
                    bool literal = task ? ProjectMpxFields.Tasks.Any(f => f.Id == id && f.IsText) : ProjectMpxFields.Resources.Any(f => f.Id == id && f.IsText);
                    literal |= !task && id == 48;
                    literal |= ProjectMpxFields.CustomMappings(task).Any(f => f.Id == id && f.Kind == "Text");
                    if (!literal) continue;
                }
                values.Add(id, text);
            }
            Tail(r, fields.Length + 1); return values;
        }
        private void Apply<T>(T entity, Dictionary<int, string> values, IEnumerable<ProjectMpxFields.Field<T>> fields) {
            foreach (var field in fields) if (values.TryGetValue(field.Id, out string? value)) { field.Read(entity, value, _values); values.Remove(field.Id); }
        }
        private int Identity(Dictionary<int, string> values, int uidField, int fallback) {
            int uid = values.TryGetValue(uidField, out var text) ? ProjectMpxValues.Integer(text) : fallback;
            values.Remove(uidField); if (uid < 0) throw new InvalidDataException("MPX identities must not be negative."); return uid;
        }
        private void Remaining(Dictionary<int, string> values) {
            if (values.Count != 0) Opaque("MPX fields " + string.Join(", ", values.Keys.OrderBy(v => v)) + " remain in the original bytes without typed mappings.");
        }
        private void ReadResource(string[] r) {
            Budget(); if (Document.Resources.Count >= 9999) throw new InvalidDataException("MPX exceeds 9999 resources.");
            var values = Values(r, _resourceFields, false);
            int row = values.TryGetValue(40, out var text) ? ProjectMpxValues.Integer(text) : Document.Resources.Count + 1;
            int uid = Identity(values, 49, row);
            if (row < 0 || row > 9999 || Document.ResourceIndex.ContainsKey(uid) || _resourceRows.ContainsKey(row)) throw new InvalidDataException("Duplicate or invalid MPX resource identity.");
            var resource = new ProjectResource(Document, uid) { DisplayId = row, Type = ProjectResourceType.Work };
            Apply(resource, values, ProjectMpxFields.Resources);
            if (values.TryGetValue(48, out string? calendar)) { resource.Calendar = FindCalendar(calendar); values.Remove(48); }
            Baseline(resource.Baselines, values, false);
            Custom(resource.CustomFields, values, false);
            Remaining(values);
            Document.Resources.Items.Add(resource); Document.ResourceIndex.Add(uid, resource); _resourceRows.Add(row, resource);
            _resource = resource; _calendar = null;
        }
        private void ReadTask(string[] r) {
            Budget(); if (Document.TaskIndex.Count >= Math.Min(9999, _options.MaxTasks)) throw new InvalidDataException("MPX exceeds its task limit.");
            var values = Values(r, _taskFields, true);
            int row = values.TryGetValue(90, out var text) ? ProjectMpxValues.Integer(text) : Document.TaskIndex.Count + 1;
            int uid = Identity(values, 98, row);
            if (row < 0 || row > 9999 || Document.TaskIndex.ContainsKey(uid) || _taskRows.ContainsKey(row)) throw new InvalidDataException("Duplicate or invalid MPX task identity.");
            int level = values.TryGetValue(3, out text) ? ProjectMpxValues.Integer(text) : row == 0 ? 0 : 1; values.Remove(3);
            if (level < 0 || level > _options.MaxOutlineDepth) throw new InvalidDataException("MPX exceeds MaxOutlineDepth.");
            if ((level == 0) != (uid == 0)) throw new InvalidDataException("MPX outline level zero is reserved for task UID zero.");
            var task = new ProjectTask(Document, uid) { DisplayId = row, SourceOutlineLevel = level };
            Apply(task, values, ProjectMpxFields.Tasks);
            if (values.TryGetValue(120, out text)) { task.SourceSummary = _values.Flag(text); values.Remove(120); }
            foreach (int field in new[] { 70, 71, 74, 75 }) if (values.TryGetValue(field, out text)) {
                _links.Add((task, text, field >= 74, field == 71 || field == 75)); values.Remove(field);
            }
            Baseline(task.Baselines, values, true); Custom(task.CustomFields, values, true); Remaining(values);
            while (_parents.Count != 0 && _parents.Peek().SourceOutlineLevel >= level) _parents.Pop();
            var parent = _parents.Count != 0 && _parents.Peek().SourceOutlineLevel > 0 ? _parents.Peek() : null;
            if (level > 1 && parent?.SourceOutlineLevel != level - 1) throw new InvalidDataException("MPX outline skips a parent level.");
            task.Parent = parent; (parent?.Children ?? Document.Tasks).Items.Add(task); _parents.Push(task);
            Document.TaskIndex.Add(uid, task); _taskRows.Add(row, task); _task = task; _taskAssignmentCount = 0;
        }
        private void Baseline(ProjectCollection<ProjectBaseline> baselines, Dictionary<int, string> values, bool task) {
            if (!values.Keys.Any(k => k == 21 || k == 31 || task && (k == 41 || k == 56 || k == 57))) return;
            var baseline = baselines.Add(); baseline.Number = 0;
            if (values.TryGetValue(21, out var text)) { baseline.Work = _values.Work(text); values.Remove(21); }
            if (values.TryGetValue(31, out text)) { baseline.Cost = _values.Number(text); values.Remove(31); }
            if (!task) return;
            if (values.TryGetValue(41, out text)) { baseline.Duration = _values.Duration(text); values.Remove(41); }
            if (values.TryGetValue(56, out text)) { baseline.Start = _values.Date(text); values.Remove(56); }
            if (values.TryGetValue(57, out text)) { baseline.Finish = _values.Date(text); values.Remove(57); }
        }
        private void Custom(ProjectCollection<ProjectCustomFieldValue> fields, Dictionary<int, string> values, bool task) {
            foreach (var mapping in ProjectMpxFields.CustomMappings(task)) {
                if (!values.TryGetValue(mapping.Id, out var text)) continue;
                var value = fields.Add(); value.FieldId = mapping.FieldId; value.Value = mapping.Parse(text, _values); values.Remove(mapping.Id);
                if (mapping.Kind == "Duration") { var duration = _values.Duration(text); value.DurationFormat = 3 + (int)duration.Unit * 2 + (duration.IsElapsed ? 1 : 0) + (duration.IsEstimated ? 32 : 0); }
                if (!Document.CustomFields.Any(f => f.FieldId == mapping.FieldId)) { var definition = Document.CustomFields.Add(); definition.FieldId = mapping.FieldId; definition.FieldName = mapping.Name; }
            }
        }
        private void Assignment(string[] r) {
            Budget();
            if (_taskAssignmentCount >= 100) throw new InvalidDataException("MPX exceeds 100 assignments per task.");
            ProjectResource? resource;
            if (Has(r, 13)) { int uid = ProjectMpxValues.Integer(r[13]); Document.ResourceIndex.TryGetValue(uid, out resource); }
            else { int row = ProjectMpxValues.Integer(Get(r, 1)); _resourceRows.TryGetValue(row, out resource); }
            if (resource == null) throw new InvalidDataException("MPX assignment refers to an undefined resource.");
            if (Document.AssignmentPairs.Contains((_task!.Uid, resource.Uid))) throw new InvalidDataException("Duplicate MPX task-resource assignment.");
            bool hasUnits = Has(r, 2);
            var item = Document.Assignments.Add(_task!, resource, hasUnits ? _values.Units(r[2]) : (ProjectUnits?)null);
            if (!hasUnits) item.Units = null;
            _taskAssignmentCount++;
            if (Has(r, 3)) item.Work = _values.Work(r[3]);
            if (Has(r, 5)) item.ActualWork = _values.Work(r[5]);
            if (Has(r, 6)) item.OvertimeWork = _values.Work(r[6]);
            if (Has(r, 7)) item.Cost = _values.Number(r[7]);
            if (Has(r, 9)) item.ActualCost = _values.Number(r[9]);
            item.Start = Has(r, 10) ? _values.Date(r[10]) : (DateTime?)null;
            item.Finish = Has(r, 11) ? _values.Date(r[11]) : (DateTime?)null;
            if (Has(r, 4) || Has(r, 8)) { var baseline = item.Baselines.Add(); baseline.Number = 0; if (Has(r, 4)) baseline.Work = _values.Work(r[4]); if (Has(r, 8)) baseline.Cost = _values.Number(r[8]); }
            if (Has(r, 12)) Opaque("Assignment delay remains in the original MPX bytes.", "/Assignment[UID=" + item.Uid + "]/Delay");
            Tail(r, 14);
        }
        private void ResolveLinks() {
            if (ProjectMpxRecords.HasAmbiguousDependencySeparator(_listSeparator) || _values.DecimalSeparator == _listSeparator.ToString()) {
                foreach (var link in _links) {
                    _token.ThrowIfCancellationRequested();
                    Opaque("Dependency syntax with this list separator is ambiguous and remains in the original bytes.", "/Task[UID=" + link.Task.Uid + "]/Dependencies");
                }
                return;
            }
            var resolved = new Dictionary<(int, int), ProjectDependency>();
            foreach (var link in _links) foreach (string piece in link.Text.Split(new[] { _listSeparator }, StringSplitOptions.RemoveEmptyEntries)) {
                _token.ThrowIfCancellationRequested();
                var match = Regex.Match(piece, @"^\s*(\d+)\s*(FS|FF|SS|SF)?\s*([+-].*)?\s*$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
                if (!match.Success) { Opaque("External or unsupported predecessor syntax remains inert: " + piece, "/Task[UID=" + link.Task.Uid + "]/Dependencies"); continue; }
                int id = ProjectMpxValues.Integer(match.Groups[1].Value);
                var lookup = link.Uids ? Document.TaskIndex : _taskRows;
                if (!lookup.TryGetValue(id, out var other)) throw new InvalidDataException("MPX dependency refers to an undefined task: " + id);
                var predecessor = link.Successors ? link.Task : other; var successor = link.Successors ? other : link.Task;
                if (predecessor == successor) throw new InvalidDataException("An MPX task cannot depend on itself.");
                ProjectDependencyType? type = match.Groups[2].Value.ToUpperInvariant() switch { "FF" => ProjectDependencyType.FinishToFinish, "SS" => ProjectDependencyType.StartToStart, "SF" => ProjectDependencyType.StartToFinish, "FS" => ProjectDependencyType.FinishToStart, _ => null };
                ProjectDuration? lag = null; decimal? percent = null; string lagText = match.Groups[3].Value;
                if (lagText.Length != 0) { if (lagText.EndsWith("%", StringComparison.Ordinal)) percent = _values.Number(lagText.TrimEnd('%')); else lag = _values.Duration(lagText); }
                var key = (predecessor.Uid, successor.Uid);
                resolved.TryGetValue(key, out var existing);
                if (existing != null) { if (existing.Type != type || !Equals(existing.Lag, lag) || existing.LagPercent != percent) throw new InvalidDataException("Conflicting MPX dependency declarations."); continue; }
                var dependency = Document.Dependencies.Add(predecessor, successor); dependency.Type = type; dependency.Lag = lag; dependency.LagPercent = percent;
                resolved.Add(key, dependency);
            }
        }
    }
}
