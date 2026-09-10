namespace OfficeIMO.Project;

internal static partial class ProjectXmlFields {
    internal static readonly ProjectXmlField<ProjectSettings>[] Settings = {
        new ProjectXmlField<ProjectSettings>("StartDate", (m, d) => ProjectXmlValue.Date(m.StartDate), (m, v, d, e) => m.StartDate = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectSettings>("FinishDate", (m, d) => ProjectXmlValue.Date(m.FinishDate), (m, v, d, e) => m.FinishDate = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectSettings>("ScheduleFromStart", (m, d) => ProjectXmlValue.Boolean(m.ScheduleFromStart), (m, v, d, e) => m.ScheduleFromStart = ProjectXmlValue.ParseBool(v)),
        new ProjectXmlField<ProjectSettings>("MinutesPerDay", (m, d) => ProjectXmlValue.Integer(m.MinutesPerDay), (m, v, d, e) => m.MinutesPerDay = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectSettings>("MinutesPerWeek", (m, d) => ProjectXmlValue.Integer(m.MinutesPerWeek), (m, v, d, e) => m.MinutesPerWeek = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectSettings>("DaysPerMonth", (m, d) => ProjectXmlValue.Integer(m.DaysPerMonth), (m, v, d, e) => m.DaysPerMonth = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectSettings>("DefaultStartTime", (m, d) => ProjectXmlValue.Clock(m.DefaultStartTime), (m, v, d, e) => m.DefaultStartTime = ProjectXmlValue.ParseClock(v)),
        new ProjectXmlField<ProjectSettings>("DefaultFinishTime", (m, d) => ProjectXmlValue.Clock(m.DefaultFinishTime), (m, v, d, e) => m.DefaultFinishTime = ProjectXmlValue.ParseClock(v)),
        new ProjectXmlField<ProjectSettings>("CurrencyCode", (m, d) => ProjectXmlValue.Text(m.CurrencyCode), (m, v, d, e) => m.CurrencyCode = v),
        new ProjectXmlField<ProjectSettings>("CurrencySymbol", (m, d) => ProjectXmlValue.Text(m.CurrencySymbol), (m, v, d, e) => m.CurrencySymbol = v),
        new ProjectXmlField<ProjectSettings>("CurrencyDigits", (m, d) => ProjectXmlValue.Integer(m.CurrencyDigits), (m, v, d, e) => m.CurrencyDigits = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectSettings>("StatusDate", (m, d) => ProjectXmlValue.Date(m.StatusDate), (m, v, d, e) => m.StatusDate = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectSettings>("DefaultTaskType", (m, d) => m.DefaultTaskType.HasValue ? ((int)m.DefaultTaskType.Value).ToString(System.Globalization.CultureInfo.InvariantCulture) : null, (m, v, d, e) => m.DefaultTaskType = (ProjectTaskType)ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectSettings>("NewTasksAreManual", (m, d) => ProjectXmlValue.Boolean(m.NewTasksAreManual), (m, v, d, e) => m.NewTasksAreManual = ProjectXmlValue.ParseBool(v)),
    };
}
