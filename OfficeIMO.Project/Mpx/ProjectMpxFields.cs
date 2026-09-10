namespace OfficeIMO.Project;

/// <summary>Typed MPX field identifiers from Microsoft's MPX 4.0 exchange contract.</summary>
internal static partial class ProjectMpxFields {
    internal sealed class Field<T> {
        internal readonly int Id;
        internal readonly string Name, ModelKey;
        internal readonly Action<T, string, ProjectMpxValues> Read;
        internal readonly Func<T, string> Write;
        internal bool IsText => ModelKey == "Name" || ModelKey == "Wbs" || ModelKey == "Notes" || ModelKey == "Contact" ||
            ModelKey == "Initials" || ModelKey == "Group" || ModelKey == "EmailAddress";
        internal Field(int id, string name, string modelKey, Action<T, string, ProjectMpxValues> read, Func<T, string> write) {
            Id = id; Name = name; ModelKey = modelKey; Read = read; Write = write;
        }
    }
    internal static string Notes(string? value) => (value ?? "").Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "\u007f");
    internal static readonly Field<ProjectTask>[] Tasks = {
        new Field<ProjectTask>(90, "ID", "DisplayId", (t, s, v) => t.DisplayId = ProjectMpxValues.Integer(s), t => ProjectMpxValues.Text(t.DisplayId)),
        new Field<ProjectTask>(1, "Name", "Name", (t, s, v) => t.Name = s, t => ProjectMpxValues.Text(t.Name)),
        new Field<ProjectTask>(2, "WBS", "Wbs", (t, s, v) => t.Wbs = s, t => ProjectMpxValues.Text(t.Wbs)),
        new Field<ProjectTask>(14, "Notes", "Notes", (t, s, v) => t.Notes = s.Replace("\u007f", "\n"), t => Notes(t.Notes)),
        new Field<ProjectTask>(15, "Contact", "Contact", (t, s, v) => t.Contact = s, t => ProjectMpxValues.Text(t.Contact)),
        new Field<ProjectTask>(20, "Work", "Work", (t, s, v) => t.Work = v.Work(s), t => ProjectMpxValues.Text(t.Work)),
        new Field<ProjectTask>(22, "Actual Work", "ActualWork", (t, s, v) => t.ActualWork = v.Work(s), t => ProjectMpxValues.Text(t.ActualWork)),
        new Field<ProjectTask>(23, "Remaining Work", "RemainingWork", (t, s, v) => t.RemainingWork = v.Work(s), t => ProjectMpxValues.Text(t.RemainingWork)),
        new Field<ProjectTask>(25, "% Work Complete", "PercentWorkComplete", (t, s, v) => t.PercentWorkComplete = v.Percent(s), t => ProjectMpxValues.Text(t.PercentWorkComplete)),
        new Field<ProjectTask>(30, "Cost", "Cost", (t, s, v) => t.Cost = v.Number(s), t => ProjectMpxValues.Text(t.Cost)),
        new Field<ProjectTask>(32, "Actual Cost", "ActualCost", (t, s, v) => t.ActualCost = v.Number(s), t => ProjectMpxValues.Text(t.ActualCost)),
        new Field<ProjectTask>(33, "Remaining Cost", "RemainingCost", (t, s, v) => t.RemainingCost = v.Number(s), t => ProjectMpxValues.Text(t.RemainingCost)),
        new Field<ProjectTask>(35, "Fixed Cost", "FixedCost", (t, s, v) => t.FixedCost = v.Number(s), t => ProjectMpxValues.Text(t.FixedCost)),
        new Field<ProjectTask>(40, "Duration", "Duration", (t, s, v) => t.Duration = v.Duration(s), t => ProjectMpxValues.Text(t.Duration)),
        new Field<ProjectTask>(42, "Actual Duration", "ActualDuration", (t, s, v) => t.ActualDuration = v.Duration(s), t => ProjectMpxValues.Text(t.ActualDuration)),
        new Field<ProjectTask>(43, "Remaining Duration", "RemainingDuration", (t, s, v) => t.RemainingDuration = v.Duration(s), t => ProjectMpxValues.Text(t.RemainingDuration)),
        new Field<ProjectTask>(44, "% Complete", "PercentComplete", (t, s, v) => t.PercentComplete = v.Percent(s), t => ProjectMpxValues.Text(t.PercentComplete)),
        new Field<ProjectTask>(50, "Start", "Start", (t, s, v) => t.Start = v.Date(s), t => ProjectMpxValues.Text(t.Start)),
        new Field<ProjectTask>(51, "Finish", "Finish", (t, s, v) => t.Finish = v.Date(s), t => ProjectMpxValues.Text(t.Finish)),
        new Field<ProjectTask>(52, "Early Start", "EarlyStart", (t, s, v) => t.EarlyStart = v.Date(s), t => ProjectMpxValues.Text(t.EarlyStart)),
        new Field<ProjectTask>(53, "Early Finish", "EarlyFinish", (t, s, v) => t.EarlyFinish = v.Date(s), t => ProjectMpxValues.Text(t.EarlyFinish)),
        new Field<ProjectTask>(54, "Late Start", "LateStart", (t, s, v) => t.LateStart = v.Date(s), t => ProjectMpxValues.Text(t.LateStart)),
        new Field<ProjectTask>(55, "Late Finish", "LateFinish", (t, s, v) => t.LateFinish = v.Date(s), t => ProjectMpxValues.Text(t.LateFinish)),
        new Field<ProjectTask>(58, "Actual Start", "ActualStart", (t, s, v) => t.ActualStart = v.Date(s), t => ProjectMpxValues.Text(t.ActualStart)),
        new Field<ProjectTask>(59, "Actual Finish", "ActualFinish", (t, s, v) => t.ActualFinish = v.Date(s), t => ProjectMpxValues.Text(t.ActualFinish)),
        new Field<ProjectTask>(68, "Constraint Date", "ConstraintDate", (t, s, v) => t.ConstraintDate = v.Date(s), t => ProjectMpxValues.Text(t.ConstraintDate)),
        new Field<ProjectTask>(81, "Milestone", "IsMilestone", (t, s, v) => t.IsMilestone = v.Flag(s), t => ProjectMpxValues.Text(t.IsMilestone)),
        new Field<ProjectTask>(82, "Critical", "IsCritical", (t, s, v) => t.IsCritical = v.Flag(s), t => ProjectMpxValues.Text(t.IsCritical)),
        new Field<ProjectTask>(91, "Constraint Type", "ConstraintType", (t, s, v) => t.ConstraintType = Constraint(s), t => ConstraintText(t.ConstraintType)),
        new Field<ProjectTask>(95, "Priority", "Priority", (t, s, v) => t.Priority = Priority(s), t => PriorityText(t.Priority)),
        new Field<ProjectTask>(80, "Fixed", "Type", (t, s, v) => t.Type = v.Flag(s) ? ProjectTaskType.FixedDuration : ProjectTaskType.FixedUnits, t => t.Type == null ? "" : ProjectMpxValues.Text(t.Type == ProjectTaskType.FixedDuration)),
        new Field<ProjectTask>(93, "Free Slack", "FreeSlackMinutes", (t, s, v) => t.FreeSlackMinutes = v.Minutes(v.Duration(s)), t => t.FreeSlackMinutes == null ? "" : ProjectMpxValues.Text(t.FreeSlackMinutes) + "m"),
        new Field<ProjectTask>(94, "Total Slack", "TotalSlackMinutes", (t, s, v) => t.TotalSlackMinutes = v.Minutes(v.Duration(s)), t => t.TotalSlackMinutes == null ? "" : ProjectMpxValues.Text(t.TotalSlackMinutes) + "m"),
    };
    internal static readonly Field<ProjectResource>[] Resources = {
        new Field<ProjectResource>(40, "ID", "DisplayId", (t, s, v) => t.DisplayId = ProjectMpxValues.Integer(s), t => ProjectMpxValues.Text(t.DisplayId)),
        new Field<ProjectResource>(1, "Name", "Name", (t, s, v) => t.Name = s, t => ProjectMpxValues.Text(t.Name)),
        new Field<ProjectResource>(2, "Initials", "Initials", (t, s, v) => t.Initials = s, t => ProjectMpxValues.Text(t.Initials)),
        new Field<ProjectResource>(3, "Group", "Group", (t, s, v) => t.Group = s, t => ProjectMpxValues.Text(t.Group)),
        new Field<ProjectResource>(10, "Notes", "Notes", (t, s, v) => t.Notes = s.Replace("\u007f", "\n"), t => Notes(t.Notes)),
        new Field<ProjectResource>(11, "Email Address", "EmailAddress", (t, s, v) => t.EmailAddress = s, t => ProjectMpxValues.Text(t.EmailAddress)),
        new Field<ProjectResource>(20, "Work", "Work", (t, s, v) => t.Work = v.Work(s), t => ProjectMpxValues.Text(t.Work)),
        new Field<ProjectResource>(22, "Actual Work", "ActualWork", (t, s, v) => t.ActualWork = v.Work(s), t => ProjectMpxValues.Text(t.ActualWork)),
        new Field<ProjectResource>(23, "Remaining Work", "RemainingWork", (t, s, v) => t.RemainingWork = v.Work(s), t => ProjectMpxValues.Text(t.RemainingWork)),
        new Field<ProjectResource>(30, "Cost", "Cost", (t, s, v) => t.Cost = v.Number(s), t => ProjectMpxValues.Text(t.Cost)),
        new Field<ProjectResource>(32, "Actual Cost", "ActualCost", (t, s, v) => t.ActualCost = v.Number(s), t => ProjectMpxValues.Text(t.ActualCost)),
        new Field<ProjectResource>(41, "Max Units", "MaxUnits", (t, s, v) => t.MaxUnits = v.Units(s), t => ProjectMpxValues.Text(t.MaxUnits)),
        new Field<ProjectResource>(42, "Standard Rate", "StandardRate", (t, s, v) => t.StandardRate = v.Rate(s), t => ProjectMpxValues.Text(t.StandardRate) + (t.StandardRate == null ? "" : "/h")),
        new Field<ProjectResource>(43, "Overtime Rate", "OvertimeRate", (t, s, v) => t.OvertimeRate = v.Rate(s), t => ProjectMpxValues.Text(t.OvertimeRate) + (t.OvertimeRate == null ? "" : "/h")),
        new Field<ProjectResource>(44, "Cost Per Use", "CostPerUse", (t, s, v) => t.CostPerUse = v.Number(s), t => ProjectMpxValues.Text(t.CostPerUse)),
    };
    private static readonly string[] Priorities = { "Lowest", "Very Low", "Lower", "Low", "Medium", "High", "Higher", "Very High", "Highest", "Do Not Level" };
    private static int Priority(string text) {
        int index = Array.FindIndex(Priorities, v => string.Equals(v, text, StringComparison.OrdinalIgnoreCase));
        if (index < 0) throw new InvalidDataException("Unknown MPX priority class: " + text);
        return (index + 1) * 100;
    }
    internal static string PriorityText(int? value) => value == null ? "" : Priorities[Math.Max(0, Math.Min(9, (int)decimal.Round(value.Value / 100m, 0, MidpointRounding.AwayFromZero) - 1))];
    private static readonly string[] Constraints = { "As Soon As Possible", "As Late As Possible", "Must Start On", "Must Finish On", "Start No Earlier Than", "Start No Later Than", "Finish No Earlier Than", "Finish No Later Than" };
    private static readonly string[] ConstraintAbbreviations = { "ASAP", "ALAP", "MSO", "MFO", "SNET", "SNLT", "FNET", "FNLT" };
    private static ProjectConstraintType Constraint(string text) {
        int index = Array.FindIndex(Constraints, v => string.Equals(v, text, StringComparison.OrdinalIgnoreCase));
        if (index < 0) index = Array.FindIndex(ConstraintAbbreviations, v => string.Equals(v, text, StringComparison.OrdinalIgnoreCase));
        if (index < 0 && int.TryParse(text, out int raw) && raw >= 0 && raw < 8) index = raw;
        if (index < 0) throw new InvalidDataException("Unknown MPX constraint: " + text);
        return (ProjectConstraintType)index;
    }
    private static string ConstraintText(ProjectConstraintType? type) => type == null ? "" : Constraints[(int)type.Value];
}
