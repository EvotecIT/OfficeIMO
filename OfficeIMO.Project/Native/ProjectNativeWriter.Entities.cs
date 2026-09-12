namespace OfficeIMO.Project;

internal sealed partial class ProjectNativeWriter {
    private void WriteTasks() {
        if (!_new && !ChangedTree("/Task")) return;
        using var editor = Editor("Task", 0x14, 0x0b400056);
        if (_new) {
            for (int i = 0; i < 3; i++) editor.AddReserved(i);
            if (!_document.TaskIndex.ContainsKey(0)) {
                editor.Add(0); editor.Integer(0, 0x0b400056, 0); editor.Integer(0, 0x0b400017, 0); editor.Integer(0, 0x0b4000f9, 0);
                editor.Set(0, 0x0b40000e, Text(_document.Name ?? "Project")); editor.Set(0, 0x0b40005c, new byte[] { 1 });
                Identity(editor, 0, 0x0b400477, NativeGuid(0, SyntheticProjectSummaryKind));
                SortPosition(editor, 0, 0x0b400479, 1);
            }
        }
        Removed(editor, "Task", _document.TaskIndex.Keys);
        int row = 0; var levels = new Dictionary<ProjectTask, int>();
        foreach (var task in _document.AllTasks) {
            _token.ThrowIfCancellationRequested(); string path = Path(task, "Task"); bool added = !editor.Contains(task.Uid);
            if (added) {
                editor.Add(task.Uid); editor.Integer(task.Uid, 0x0b400056, task.Uid);
                Identity(editor, task.Uid, 0x0b400477, EntityGuid(task, 1));
                editor.Integer(task.Uid, 0x0b400019, _profile == ProjectNativeProfile.Mpp8 ? 4 : 500); editor.Integer(task.Uid, 0x0b400080, 0); editor.Integer(task.Uid, 0x0b400011, 0);
                if (_profile == ProjectNativeProfile.Mpp14) {
                    editor.Set(task.Uid, 0x0b4004ff, new byte[] { 1 }); editor.Set(task.Uid, 0x0b400500, new byte[] { 0 });
                }
                editor.Integer(task.Uid, 0x0b4000b5, task.Duration.HasValue ? DurationFormat(task.Duration.Value) : 7);
                if (task.Duration.HasValue && !task.RemainingDuration.HasValue && !task.ActualDuration.HasValue && (task.PercentComplete ?? 0) == 0)
                    editor.Integer(task.Uid, 0x0b40001f, Exact(Minutes(task.Duration.Value) * 10));
            }
            int level = task.Uid == 0 ? 0 : task.Parent == null ? 1 : levels[task.Parent] + 1; levels[task] = level;
            WriteTaskFields(editor, task.Uid, path);
            if (task.IsNull != true) Handle(path + "/IsNull");
            Handle(path + "/Uid"); Handle(path + "/Parent"); Handle(path + "/Position"); Handle(path + "/IsSummary");
            int displayId = task.Uid == 0 ? 0 : ++row;
            CheckDisplayId(task.DisplayId, displayId, path);
            if (_new || StructureChanged || added || Changed(path + "/DisplayId")) {
                editor.Integer(task.Uid, 0x0b400017, displayId);
                SortPosition(editor, task.Uid, 0x0b400479, displayId + 1);
                editor.Integer(task.Uid, 0x0b4000f9, level); editor.Integer(task.Uid, 0x0b4000a0, task.Parent?.Uid ?? 0);
                editor.Set(task.Uid, 0x0b40005c, new byte[] { task.IsSummary ? (byte)1 : (byte)0 });
            }
            if (_new || StructureChanged || added || _taskGuidsChanged) {
                var parent = task.Parent;
                Guid parentGuid = parent != null ? EntityGuid(parent, 1) : task.Uid == 0 ? Guid.Empty
                    : _document.TaskIndex.TryGetValue(0, out var projectSummary) ? EntityGuid(projectSummary, 1) : NativeGuid(0, SyntheticProjectSummaryKind);
                Identity(editor, task.Uid, 0x0b40047f, parentGuid);
            }
            if (!ProjectXmlValue.TryTaskDurationFormat(task, out _))
                AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_DURATION_FORMAT", ProjectDiagnosticSeverity.Error,
                    "Native task duration, actual duration, and remaining duration share one format. Use the same unit and flags.", path));
            WriteBaselines(editor, task.Uid, path, task.Baselines, 0);
            WriteCustomFields(editor, task.Uid, path, task.CustomFields, true);
        }
        editor.Export(_replacements);
    }
    private void WriteResources() {
        if (!_new && !ChangedTree("/Resource") && !_calendarGuidsChanged) return;
        using var editor = Editor("Rsc", 0x15, 0x0c40001b);
        if (_new) {
            for (int i = 0; i < 3; i++) editor.AddReserved(i);
            if (!_document.ResourceIndex.ContainsKey(0)) {
                editor.Add(0); editor.Integer(0, 0x0c40001b, 0); editor.Integer(0, 0x0c400000, 0);
                if (_profile != ProjectNativeProfile.Mpp8) editor.Set(0, 0x0c40012a, new byte[] { 1 });
            }
        }
        Removed(editor, "Resource", _document.ResourceIndex.Keys);
        int row = 0;
        foreach (var resource in _document.Resources) {
            _token.ThrowIfCancellationRequested(); string path = Path(resource, "Resource"); bool added = !editor.Contains(resource.Uid);
            int displayId = resource.Uid == 0 ? 0 : ++row;
            CheckDisplayId(resource.DisplayId, displayId, path);
            if (added) { editor.Add(resource.Uid); editor.Integer(resource.Uid, 0x0c40001b, resource.Uid);
                Identity(editor, resource.Uid, 0x0c4002d8, EntityGuid(resource, 2)); }
            if (added || StructureChanged || Changed(path + "/DisplayId")) {
                editor.Integer(resource.Uid, 0x0c400000, displayId);
                SortPosition(editor, resource.Uid, 0x0c4002da, displayId + 1);
            }
            WriteResourceFields(editor, resource.Uid, path); Handle(path + "/Uid"); Handle(path + "/Type");
            if (resource.IsNull != true) Handle(path + "/IsNull");
            if (resource.Calendar?.IsBaseCalendar == true) AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_RESOURCE_CALENDAR", ProjectDiagnosticSeverity.Error,
                "Native resources require a resource-specific derived calendar. Create it with Calendars.Add(resourceName, baseCalendar).", path + "/Calendar"));
            if (added || Changed(path + "/Calendar") || _calendarGuidsChanged)
                Identity(editor, resource.Uid, 0x0c4002d9, resource.Calendar == null ? Guid.Empty : EntityGuid(resource.Calendar, 5));
            if (added || Changed(path + "/Type")) {
                if (_profile != ProjectNativeProfile.Mpp8) editor.Set(resource.Uid, 0x0c40012a, new byte[] { resource.Type == null || resource.Type == ProjectResourceType.Work ? (byte)1 : (byte)0 });
                else if (resource.Type != null && resource.Type != ProjectResourceType.Work) AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_RESOURCE_TYPE", ProjectDiagnosticSeverity.Error,
                    "Project 98 output supports work resources only.", path + "/Type"));
                if (_profile.HasExtendedRecords) editor.Set(resource.Uid, 0x0c4002df, new byte[] { resource.Type == ProjectResourceType.Cost ? (byte)1 : (byte)0 });
                else if (resource.Type == ProjectResourceType.Cost) AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_RESOURCE_TYPE", ProjectDiagnosticSeverity.Error,
                    "Cost resources require MPP12 or later.", path + "/Type"));
            }
            WriteBaselines(editor, resource.Uid, path, resource.Baselines, 1);
            WriteCustomFields(editor, resource.Uid, path, resource.CustomFields, false);
        }
        editor.Export(_replacements);
    }
    private void CheckDisplayId(int? requested, int actual, string path) {
        string key = path + "/DisplayId"; Handle(key);
        if (Changed(key) && requested != actual) AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_DISPLAY_ORDER", ProjectDiagnosticSeverity.Error,
            "Native display IDs follow collection order. Move tasks through MoveTo; do not assign a conflicting or absent display ID.", key));
    }
    private void WriteAssignments() {
        if (!_new && !ChangedTree("/Assignment") && !_taskGuidsChanged && !_resourceGuidsChanged) return;
        using var editor = Editor("Assn", 0x17, 0x0f400000);
        Removed(editor, "Assignment", _document.AssignmentIndex.Keys);
        foreach (var assignment in _document.Assignments) {
            _token.ThrowIfCancellationRequested(); string path = Path(assignment, "Assignment");
            if (!editor.Contains(assignment.Uid)) {
                editor.Add(assignment.Uid); editor.Integer(assignment.Uid, 0x0f400000, assignment.Uid);
                Identity(editor, assignment.Uid, 0x0f40027c, EntityGuid(assignment, 3));
                if (assignment.Task != null) Identity(editor, assignment.Uid, 0x0f40027d, EntityGuid(assignment.Task, 1));
                if (assignment.Resource != null) Identity(editor, assignment.Uid, 0x0f40027e, EntityGuid(assignment.Resource, 2));
                if (_profile.HasExtendedRecords) editor.Set(assignment.Uid, 0x0f400283, new byte[] { 1 });
                if (_profile != ProjectNativeProfile.Mpp8) { editor.Set(assignment.Uid, 0x0f400118, new byte[] { 1 }); editor.Set(assignment.Uid, 0x0f40010d, new byte[] { 1 }); }
                editor.Integer(assignment.Uid, 0x0f400037, 7);
                if (assignment.Work.HasValue && !assignment.RemainingWork.HasValue && !assignment.ActualWork.HasValue) {
                    editor.Set(assignment.Uid, 0x0f40000c, BitConverter.GetBytes((double)(assignment.Work.Value.Minutes * 1000)));
                    editor.Set(assignment.Uid, 0x0f40000b, BitConverter.GetBytes((double)(assignment.Work.Value.Minutes * 1000)));
                }
            }
            Handle(path + "/Uid"); WriteAssignmentFields(editor, assignment.Uid, path);
            if ((_taskGuidsChanged || Changed(path + "/Task")) && assignment.Task != null) Identity(editor, assignment.Uid, 0x0f40027d, EntityGuid(assignment.Task, 1));
            if ((_resourceGuidsChanged || Changed(path + "/Resource")) && assignment.Resource != null) Identity(editor, assignment.Uid, 0x0f40027e, EntityGuid(assignment.Resource, 2));
            WriteBaselines(editor, assignment.Uid, path, assignment.Baselines, 2);
        }
        editor.Export(_replacements);
    }
    private void WriteDependencies() {
        // Dependency IDs are container-local. Keep all untouched dependency records verbatim.
        if (!_new && !ChangedTree("/Dependency") && !_taskGuidsChanged) return;
        Handle("/Dependency/Count");
        using var editor = Editor("Cons", 0x18, 0x0e400000);
        foreach (int uid in editor.Uids.ToArray()) editor.Delete(uid);
        int index = 0;
        foreach (var link in _document.Dependencies) {
            _token.ThrowIfCancellationRequested(); int uid = index + 1; string path = "/Dependency[" + index++ + "]";
            editor.Add(uid); editor.Integer(uid, 0x0e400000, uid);
            Identity(editor, uid, 0x0e400015, NativeGuid(uid, 4));
            editor.Integer(uid, 0x0e400002, link.Predecessor?.Uid ?? link.SourcePredecessorUid); editor.Integer(uid, 0x0e400005, link.Successor.Uid);
            editor.Integer(uid, 0x0e400007, (int)(link.Type ?? ProjectDependencyType.FinishToStart));
            editor.Integer(uid, 0x0e400009, link.LagPercent.HasValue ? Exact(link.LagPercent.Value) : link.Lag.HasValue ? Exact(Minutes(link.Lag.Value) * 10) : 0);
            editor.Integer(uid, 0x0e40000a, link.LagPercent.HasValue ? link.PercentageLagFormat : link.Lag.HasValue ? DurationFormat(link.Lag.Value) : 7);
            if (link.Predecessor != null) Identity(editor, uid, 0x0e400016, EntityGuid(link.Predecessor, 1));
            Identity(editor, uid, 0x0e400017, EntityGuid(link.Successor, 1));
            if (_profile == ProjectNativeProfile.Mpp14) editor.Set(uid, 0x0e40001c, new byte[] { 1 });
            foreach (string name in new[] { "Predecessor", "Successor", "Type", "Lag", "LagPercent", "LagPercentIsElapsed", "LagPercentIsEstimated" }) Handle(path + "/" + name);
        }
        foreach (var old in Original("/Dependency").Where(p => !_current.ContainsKey(p.Key))) Handle(old.Key);
        editor.Export(_replacements);
    }
}
