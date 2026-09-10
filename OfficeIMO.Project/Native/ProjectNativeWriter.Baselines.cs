namespace OfficeIMO.Project;

internal sealed partial class ProjectNativeWriter {
    private static readonly uint[] TaskBaselineStarts = { 0x1e2, 0x1ed, 0x1f8, 0x203, 0x20e, 0x220, 0x22b, 0x236, 0x241, 0x24c };
    private void WriteBaselines(IProjectNativeTableEditor editor, int uid, string parent, ProjectCollection<ProjectBaseline> baselines, int owner) {
        string prefix = parent + "/Baseline";
        if (!ChangedTree(prefix)) return;
        Handle(prefix + "/Count");
        (uint Work, uint Cost, uint Start, uint Finish, uint Duration, uint Format) Ids(int number) {
            if (owner == 0) {
                uint first = number == 0 ? 0 : 0x0b400000 | TaskBaselineStarts[number - 1];
                return number == 0 ? (0x0b400001, 0x0b400006, 0x0b40002b, 0x0b40002c, 0x0b40001b, 0x0b4000b3)
                    : (first + 3, first + 2, first, first + 1, first + 5, first + 6);
            }
            if (owner == 1) { uint work = number == 0 ? 0x0c40000fu : 0x0c400156u + (uint)(number - 1) * 10; return (work, number == 0 ? 0x0c400011u : work + 1, 0, 0, 0, 0); }
            uint assignmentWork = number == 0 ? 0x0f400010u : 0x0f400121u + (uint)(number - 1) * 9;
            return (assignmentWork, number == 0 ? 0x0f400020u : assignmentWork + 1, number == 0 ? 0x0f400092u : assignmentWork + 6,
                number == 0 ? 0x0f400093u : assignmentWork + 7, 0, 0);
        }
        foreach (var old in Original(prefix).Where(p => p.Key.EndsWith("/Number", StringComparison.Ordinal))) {
            if (!(old.Value is int number) || number < 0 || number > 10) continue;
            var ids = Ids(number);
            foreach (uint id in new[] { ids.Work, ids.Cost, ids.Start, ids.Finish, ids.Duration, ids.Format }) if (id != 0) editor.Set(uid, id, null);
        }
        for (int index = 0; index < baselines.Count; index++) {
            _token.ThrowIfCancellationRequested(); var item = baselines[index];
            if (!item.Number.HasValue || item.Number < 0 || item.Number > 10) continue;
            string path = prefix + "[" + index + "]"; var ids = Ids(item.Number.Value); Handle(path + "/Number");
            Field(editor, uid, path, "Work", ids.Work, Kind.Work, force: true); Field(editor, uid, path, "Cost", ids.Cost, Kind.Number, 100, force: true);
            if (ids.Start != 0) { Field(editor, uid, path, "Start", ids.Start, Kind.Date, force: true); Field(editor, uid, path, "Finish", ids.Finish, Kind.Date, force: true); }
            if (ids.Duration != 0) Field(editor, uid, path, "Duration", ids.Duration, Kind.Duration, format: ids.Format, force: true);
        }
        foreach (var old in Original(prefix).Where(p => !_current.ContainsKey(p.Key))) Handle(old.Key);
    }
}
