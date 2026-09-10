namespace OfficeIMO.Project;

internal static class ProjectNativeLegacy8Values {
    /// <summary>Converts the ten Project 98 priority classes to the shared numeric scale.</summary>
    internal static int? Priority(int? value) {
        if (!value.HasValue) return null;
        if (value < 0 || value > 9) throw new InvalidDataException("Project 98 task priority is outside its ten classes.");
        return (value.Value + 1) * 100;
    }
}

internal sealed partial class ProjectNativeWriter {
    private void WritePriority8(IProjectNativeTableEditor editor, int uid, string path) {
        string key = path + "/Priority"; Handle(key);
        if (!Changed(key)) return;
        if (!_current.TryGetValue(key, out var value) || value == null) { editor.Set(uid, 0x0b400019, null); return; }
        int priority = (int)value;
        if (priority < 0 || priority > 1000) throw new ArgumentOutOfRangeException(nameof(value), "Task priority must be between zero and 1000.");
        int stored = Math.Max(0, Math.Min(9, (priority + 50) / 100 - 1));
        editor.Integer(uid, 0x0b400019, stored);
        if (ProjectNativeLegacy8Values.Priority(stored) != priority)
            Loss("PROJECT_NATIVE_PRIORITY_QUANTIZED", "Project 98 stores ten priority classes. This priority is rounded to " + (stored + 1) * 100 + ".", key);
    }
}
