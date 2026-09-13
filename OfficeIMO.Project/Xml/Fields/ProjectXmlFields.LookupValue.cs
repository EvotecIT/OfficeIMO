namespace OfficeIMO.Project;

internal static partial class ProjectXmlFields {
    internal static readonly ProjectXmlField<ProjectLookupValue>[] LookupValue = {
        new ProjectXmlField<ProjectLookupValue>("ID", (m, d) => ProjectXmlValue.Integer(m.Id), (m, v, d, e) => m.Id = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectLookupValue>("Value", (m, d) => ProjectXmlValue.Text(m.Value), (m, v, d, e) => m.Value = v),
        new ProjectXmlField<ProjectLookupValue>("Description", (m, d) => ProjectXmlValue.Text(m.Description), (m, v, d, e) => m.Description = v),
        new ProjectXmlField<ProjectLookupValue>("GUID", (m, d) => ProjectXmlValue.Text(m.Guid), (m, v, d, e) => m.Guid = v),
    };
}
