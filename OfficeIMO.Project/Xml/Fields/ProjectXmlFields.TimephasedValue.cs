namespace OfficeIMO.Project;

internal static partial class ProjectXmlFields {
    internal static readonly ProjectXmlField<ProjectTimephasedValue>[] TimephasedValue = {
        new ProjectXmlField<ProjectTimephasedValue>("Type", (m, d) => ProjectXmlValue.Integer(m.Type), (m, v, d, e) => m.Type = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectTimephasedValue>("UID", (m, d) => ProjectXmlValue.Integer(m.Uid), (m, v, d, e) => m.Uid = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectTimephasedValue>("Start", (m, d) => ProjectXmlValue.Date(m.Start), (m, v, d, e) => m.Start = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectTimephasedValue>("Finish", (m, d) => ProjectXmlValue.Date(m.Finish), (m, v, d, e) => m.Finish = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectTimephasedValue>("Unit", (m, d) => ProjectXmlValue.Integer(m.Unit), (m, v, d, e) => m.Unit = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectTimephasedValue>("Value", (m, d) => ProjectXmlValue.Text(m.Value), (m, v, d, e) => m.Value = v),
    };
}
