namespace OfficeIMO.Project;

internal static partial class ProjectXmlFields {
    internal static readonly ProjectXmlField<ProjectCustomFieldDefinition>[] CustomFieldDefinition = {
        new ProjectXmlField<ProjectCustomFieldDefinition>("Ltuid", (m, d) => m.LookupTableGuid, (m, v, d, e) => m.LookupTableGuid = v),
        new ProjectXmlField<ProjectCustomFieldDefinition>("FieldID", (m, d) => ProjectXmlValue.Text(m.FieldId), (m, v, d, e) => m.FieldId = v),
        new ProjectXmlField<ProjectCustomFieldDefinition>("FieldName", (m, d) => ProjectXmlValue.Text(m.FieldName), (m, v, d, e) => m.FieldName = v),
        new ProjectXmlField<ProjectCustomFieldDefinition>("Alias", (m, d) => ProjectXmlValue.Text(m.Alias), (m, v, d, e) => m.Alias = v),
        new ProjectXmlField<ProjectCustomFieldDefinition>("CFType", (m, d) => ProjectXmlValue.Integer(m.FieldType), (m, v, d, e) => m.FieldType = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectCustomFieldDefinition>("Formula", (m, d) => ProjectXmlValue.Text(m.Formula), (m, v, d, e) => m.Formula = v),
        new ProjectXmlField<ProjectCustomFieldDefinition>("CalculationType", (m, d) => ProjectXmlValue.Integer(m.SummaryCalculation), (m, v, d, e) => m.SummaryCalculation = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectCustomFieldDefinition>("RollupType", (m, d) => ProjectXmlValue.Integer(m.RollupType), (m, v, d, e) => m.RollupType = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectCustomFieldDefinition>("Guid", (m, d) => ProjectXmlValue.Text(m.Guid), (m, v, d, e) => m.Guid = v),
        new ProjectXmlField<ProjectCustomFieldDefinition>("RestrictValues", (m, d) => ProjectXmlValue.Boolean(m.RestrictValues), (m, v, d, e) => m.RestrictValues = ProjectXmlValue.ParseBool(v)),
    };
}
