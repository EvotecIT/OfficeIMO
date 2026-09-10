namespace OfficeIMO.Project;

internal static partial class ProjectXmlFields {
    internal static readonly ProjectXmlField<ProjectCustomFieldValue>[] CustomFieldValue = {
        new ProjectXmlField<ProjectCustomFieldValue>("FieldID", (m, d) => ProjectXmlValue.Text(m.FieldId), (m, v, d, e) => m.FieldId = v),
        new ProjectXmlField<ProjectCustomFieldValue>("Value", (m, d) => ProjectXmlValue.Text(m.Value), (m, v, d, e) => m.Value = v),
        new ProjectXmlField<ProjectCustomFieldValue>("ValueID", (m, d) => ProjectXmlValue.Text(m.ValueId), (m, v, d, e) => m.ValueId = v),
        new ProjectXmlField<ProjectCustomFieldValue>("ValueGUID", (m, d) => ProjectXmlValue.Text(m.ValueGuid), (m, v, d, e) => m.ValueGuid = v),
        new ProjectXmlField<ProjectCustomFieldValue>("DurationFormat", (m, d) => ProjectXmlValue.Integer(m.DurationFormat), (m, v, d, e) => m.DurationFormat = ProjectXmlValue.ParseInt(v)),
    };
}
