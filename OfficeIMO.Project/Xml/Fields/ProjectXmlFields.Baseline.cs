namespace OfficeIMO.Project;

internal static partial class ProjectXmlFields {
    internal static readonly ProjectXmlField<ProjectBaseline>[] Baseline = {
        new ProjectXmlField<ProjectBaseline>("Number", (m, d) => ProjectXmlValue.Integer(m.Number), (m, v, d, e) => m.Number = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectBaseline>("Start", (m, d) => ProjectXmlValue.Date(m.Start), (m, v, d, e) => m.Start = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectBaseline>("Finish", (m, d) => ProjectXmlValue.Date(m.Finish), (m, v, d, e) => m.Finish = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectBaseline>("Duration", (m, d) => ProjectXmlValue.Duration(m.Duration, d), (m, v, d, e) => m.Duration = ProjectXmlValue.ParseDuration(v, (int?)e.Element(e.Name.Namespace + "DurationFormat"), d)),
        new ProjectXmlField<ProjectBaseline>("Work", (m, d) => ProjectXmlValue.Work(m.Work), (m, v, d, e) => m.Work = ProjectXmlValue.ParseWork(v)),
        new ProjectXmlField<ProjectBaseline>("Cost", (m, d) => ProjectXmlValue.Money(m.Cost), (m, v, d, e) => m.Cost = ProjectXmlValue.ParseMoney(v)),
        new ProjectXmlField<ProjectBaseline>("FixedCost", (m, d) => ProjectXmlValue.Money(m.FixedCost), (m, v, d, e) => m.FixedCost = ProjectXmlValue.ParseMoney(v)),
        new ProjectXmlField<ProjectBaseline>("BCWS", (m, d) => ProjectXmlValue.Money(m.Bcws), (m, v, d, e) => m.Bcws = ProjectXmlValue.ParseMoney(v)),
        new ProjectXmlField<ProjectBaseline>("BCWP", (m, d) => ProjectXmlValue.Money(m.Bcwp), (m, v, d, e) => m.Bcwp = ProjectXmlValue.ParseMoney(v)),
    };
}
