namespace OfficeIMO.Project;

internal sealed partial class ProjectNativeWriter {
    private void WriteTaskFields(IProjectNativeTableEditor editor, int uid, string path) {
        Field(editor, uid, path, "Wbs", 0x0b400010, Kind.Text); Field(editor, uid, path, "Contact", 0x0b400070, Kind.Text);
        Field(editor, uid, path, "PhysicalPercentComplete", 0x0b40045f);
        Field(editor, uid, path, "PercentComplete", 0x0b400000u | 0x20, Kind.Integer, 1);
        Field(editor, uid, path, "PercentWorkComplete", 0x0b400000u | 0x21, Kind.Integer, 1);
        if (_profile == ProjectNativeProfile.Mpp8) WritePriority8(editor, uid, path);
        else Field(editor, uid, path, "Priority", 0x0b400000u | 0x19, Kind.Integer, 1);
        Field(editor, uid, path, "ConstraintType", 0x0b400000u | 0x11, Kind.Integer, 1);
        Field(editor, uid, path, "Type", 0x0b400000u | 0x80, Kind.Integer, 1);
        Field(editor, uid, path, "Calendar", 0x0b400000u | 0x191, Kind.Integer, 1);
        Field(editor, uid, path, "Name", 0x0b400000u | 0x0e, Kind.Text, 1);
        Field(editor, uid, path, "Start", 0x0b400000u | 0x23, Kind.Date, 1);
        Field(editor, uid, path, "Finish", 0x0b400000u | 0x24, Kind.Date, 1);
        Field(editor, uid, path, "EarlyStart", 0x0b400000u | 0x25, Kind.Date, 1);
        Field(editor, uid, path, "EarlyFinish", 0x0b400000u | 0x26, Kind.Date, 1);
        if (_profile == ProjectNativeProfile.Mpp8 && Changed(path + "/Start") &&
            (!_current.TryGetValue(path + "/EarlyStart", out var earlyStart) || earlyStart == null) &&
            _current.TryGetValue(path + "/Start", out var start) && start is DateTime startDate)
            editor.Set(uid, 0x0b400025, BitConverter.GetBytes(ProjectNativeCreation.Date(startDate)));
        if (_profile == ProjectNativeProfile.Mpp8 && Changed(path + "/Finish") &&
            (!_current.TryGetValue(path + "/EarlyFinish", out var earlyFinish) || earlyFinish == null) &&
            _current.TryGetValue(path + "/Finish", out var finish) && finish is DateTime finishDate)
            editor.Set(uid, 0x0b400026, BitConverter.GetBytes(ProjectNativeCreation.Date(finishDate)));
        if (_profile == ProjectNativeProfile.Mpp8 && (Changed(path + "/Finish") || Changed(path + "/EarlyFinish")) &&
            _current.TryGetValue(path + "/Finish", out var finalDate) && finalDate is DateTime &&
            _current.TryGetValue(path + "/EarlyFinish", out var earlyDate) && earlyDate is DateTime && !Equals(finalDate, earlyDate))
            AddDiagnostic(new ProjectDiagnostic("PROJECT_NATIVE_LEGACY_FINISH_CACHE", ProjectDiagnosticSeverity.Error,
                "The qualified Project 98 date profile requires Finish and EarlyFinish to agree. Independent readers use the early-finish cache for the displayed finish.", path + "/Finish"));
        Field(editor, uid, path, "LateStart", 0x0b400000u | 0x27, Kind.Date, 1);
        Field(editor, uid, path, "LateFinish", 0x0b400000u | 0x28, Kind.Date, 1);
        Field(editor, uid, path, "ActualStart", 0x0b400000u | 0x29, Kind.Date, 1);
        Field(editor, uid, path, "ActualFinish", 0x0b400000u | 0x2a, Kind.Date, 1);
        Field(editor, uid, path, "Deadline", 0x0b400000u | 0x1b5, Kind.Date, 1);
        Field(editor, uid, path, "ConstraintDate", 0x0b400000u | 0x12, Kind.Date, 1);
        if (_profile == ProjectNativeProfile.Mpp14) {
            Field(editor, uid, path, "IsManual", 0x0b400000u | 0x500, Kind.Boolean, 1);
            Field(editor, uid, path, "IsActive", 0x0b400000u | 0x4ff, Kind.Boolean, 1);
        } else {
            // Project 2007 has automatic, active tasks only. Opposite values remain
            // unhandled so the model inventory diagnoses their conversion loss.
            if (!_current.TryGetValue(path + "/IsManual", out var manual) || manual == null || Equals(manual, false)) Handle(path + "/IsManual");
            if (!_current.TryGetValue(path + "/IsActive", out var active) || active == null || Equals(active, true)) Handle(path + "/IsActive");
        }
        Field(editor, uid, path, "IsMilestone", 0x0b400000u | 0x18, Kind.Boolean, 1);
        Field(editor, uid, path, "IsCritical", 0x0b400000u | 0x13, Kind.Boolean, 1);
        Field(editor, uid, path, "EffortDriven", 0x0b400000u | 0x84, Kind.Boolean, 1);
        Field(editor, uid, path, "Guid", 0x0b400000u | 0x477, Kind.Guid, 1);
        Field(editor, uid, path, "Work", 0x0b400000u | 0x00, Kind.Work, 1);
        Field(editor, uid, path, "ActualWork", 0x0b400000u | 0x02, Kind.Work, 1);
        Field(editor, uid, path, "RemainingWork", 0x0b400000u | 0x04, Kind.Work, 1);
        Field(editor, uid, path, "Cost", 0x0b400000u | 0x05, Kind.Number, 100);
        Field(editor, uid, path, "ActualCost", 0x0b400000u | 0x07, Kind.Number, 100);
        Field(editor, uid, path, "RemainingCost", 0x0b400000u | 0x0a, Kind.Number, 100);
        Field(editor, uid, path, "FixedCost", 0x0b400000u | 0x08, Kind.Number, 100);
        Field(editor, uid, path, "FreeSlackMinutes", 0x0b400000u | 0x15, Kind.ScaledInteger, 10);
        // Total slack has no record in the qualified producer map. Let the model inventory
        // report its omission instead of inventing a field or failing inside the table editor.
        Field(editor, uid, path, "Duration", 0x0b400000u | 0x1d, Kind.Duration, 1, 0x0b400000u | 0xb5);
        Field(editor, uid, path, "ActualDuration", 0x0b400000u | 0x1c, Kind.Duration, 1, 0x0b400000u | 0xb5);
        Field(editor, uid, path, "RemainingDuration", 0x0b400000u | 0x1f, Kind.Duration, 1, 0x0b400000u | 0xb5);
    }
    private void WriteResourceFields(IProjectNativeTableEditor editor, int uid, string path) {
        Field(editor, uid, path, "Initials", 0x0c400002, Kind.Text); Field(editor, uid, path, "Group", 0x0c400003, Kind.Text); Field(editor, uid, path, "EmailAddress", 0x0c400023, Kind.Text);
        Field(editor, uid, path, "Calendar", 0x0c400000u | 0x38, Kind.Integer, 1);
        Field(editor, uid, path, "Name", 0x0c400000u | 0x01, Kind.Text, 1);
        Field(editor, uid, path, "MaterialLabel", 0x0c400000u | 0x12b, Kind.Text, 1);
        Field(editor, uid, path, "Guid", 0x0c400000u | 0x2d8, Kind.Guid, 1);
        Field(editor, uid, path, "Work", 0x0c400000u | 0x0d, Kind.Work, 1);
        Field(editor, uid, path, "ActualWork", 0x0c400000u | 0x0e, Kind.Work, 1);
        Field(editor, uid, path, "RemainingWork", 0x0c400000u | 0x16, Kind.Work, 1);
        Field(editor, uid, path, "StandardRate", 0x0c400000u | 0x06, Kind.Number, 1);
        Field(editor, uid, path, "OvertimeRate", 0x0c400000u | 0x07, Kind.Number, 1);
        Field(editor, uid, path, "CostPerUse", 0x0c400000u | 0x12, Kind.Number, 100);
        Field(editor, uid, path, "Cost", 0x0c400000u | 0x0c, Kind.Number, 100);
        Field(editor, uid, path, "ActualCost", 0x0c400000u | 0x0b, Kind.Number, 100);
        Field(editor, uid, path, "MaxUnits", 0x0c400000u | 0x04, Kind.Units, 1);
    }
    private void WriteAssignmentFields(IProjectNativeTableEditor editor, int uid, string path) {
        Field(editor, uid, path, "Task", 0x0f400000u | 0x01, Kind.Integer, 1);
        Field(editor, uid, path, "Resource", 0x0f400000u | 0x02, Kind.Integer, 1);
        Field(editor, uid, path, "Guid", 0x0f400000u | 0x27c, Kind.Guid, 1);
        Field(editor, uid, path, "Start", 0x0f400000u | 0x14, Kind.Date, 1);
        Field(editor, uid, path, "Finish", 0x0f400000u | 0x15, Kind.Date, 1);
        // Assignment actual endpoints and percent-complete caches are not present in this
        // producer map. The model inventory applies the explicit unsupported-field policy.
        Field(editor, uid, path, "Work", 0x0f400000u | 0x08, Kind.Work, 1);
        Field(editor, uid, path, "ActualWork", 0x0f400000u | 0x0a, Kind.Work, 1);
        Field(editor, uid, path, "RemainingWork", 0x0f400000u | 0x0c, Kind.Work, 1);
        Field(editor, uid, path, "OvertimeWork", 0x0f400000u | 0x09, Kind.Work, 1);
        Field(editor, uid, path, "ActualOvertimeWork", 0x0f400000u | 0x0d, Kind.Work, 1);
        Field(editor, uid, path, "Cost", 0x0f400000u | 0x1a, Kind.Number, 100);
        Field(editor, uid, path, "ActualCost", 0x0f400000u | 0x1c, Kind.Number, 100);
        Field(editor, uid, path, "RemainingCost", 0x0f400000u | 0x1d, Kind.Number, 100);
        Field(editor, uid, path, "Units", 0x0f400000u | 0x07, Kind.Units, 1);
    }
}
