namespace OfficeIMO.Project;

internal static partial class ProjectXmlFields {
    internal static readonly ProjectXmlField<ProjectOutlineCodeDefinition>[] OutlineCodeDefinition = {
        new ProjectXmlField<ProjectOutlineCodeDefinition>("Guid", (m, d) => ProjectXmlValue.Text(m.Guid), (m, v, d, e) => m.Guid = v),
        new ProjectXmlField<ProjectOutlineCodeDefinition>("FieldID", (m, d) => ProjectXmlValue.Text(m.FieldId), (m, v, d, e) => m.FieldId = v),
        new ProjectXmlField<ProjectOutlineCodeDefinition>("FieldName", (m, d) => ProjectXmlValue.Text(m.FieldName), (m, v, d, e) => m.FieldName = v),
        new ProjectXmlField<ProjectOutlineCodeDefinition>("Alias", (m, d) => ProjectXmlValue.Text(m.Alias), (m, v, d, e) => m.Alias = v),
        new ProjectXmlField<ProjectOutlineCodeDefinition>("LeafOnly", (m, d) => ProjectXmlValue.Boolean(m.LeafOnly), (m, v, d, e) => m.LeafOnly = ProjectXmlValue.ParseBool(v)),
        new ProjectXmlField<ProjectOutlineCodeDefinition>("AllLevelsRequired", (m, d) => ProjectXmlValue.Boolean(m.AllLevelsRequired), (m, v, d, e) => m.AllLevelsRequired = ProjectXmlValue.ParseBool(v)),
        new ProjectXmlField<ProjectOutlineCodeDefinition>("OnlyTableValuesAllowed", (m, d) => ProjectXmlValue.Boolean(m.OnlyTableValuesAllowed), (m, v, d, e) => m.OnlyTableValuesAllowed = ProjectXmlValue.ParseBool(v)),
    };
    internal static readonly ProjectXmlField<ProjectOutlineCodeMask>[] OutlineCodeMask = {
        new ProjectXmlField<ProjectOutlineCodeMask>("Level", (m, d) => ProjectXmlValue.Integer(m.Level), (m, v, d, e) => m.Level = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectOutlineCodeMask>("Type", (m, d) => ProjectXmlValue.Integer(m.Type), (m, v, d, e) => m.Type = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectOutlineCodeMask>("Length", (m, d) => ProjectXmlValue.Integer(m.Length), (m, v, d, e) => m.Length = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectOutlineCodeMask>("Separator", (m, d) => ProjectXmlValue.Text(m.Separator), (m, v, d, e) => m.Separator = v),
    };
    internal static readonly ProjectXmlField<ProjectOutlineCodeLookupValue>[] OutlineCodeLookupValue = {
        new ProjectXmlField<ProjectOutlineCodeLookupValue>("ValueID", (m, d) => ProjectXmlValue.Integer(m.ValueId), (m, v, d, e) => m.ValueId = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectOutlineCodeLookupValue>("FieldGUID", (m, d) => ProjectXmlValue.Text(m.Guid), (m, v, d, e) => m.Guid = v),
        new ProjectXmlField<ProjectOutlineCodeLookupValue>("ParentValueID", (m, d) => ProjectXmlValue.Integer(m.ParentValueId), (m, v, d, e) => m.ParentValueId = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectOutlineCodeLookupValue>("Type", (m, d) => ProjectXmlValue.Integer(m.Type), (m, v, d, e) => m.Type = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectOutlineCodeLookupValue>("Value", (m, d) => ProjectXmlValue.Text(m.Value), (m, v, d, e) => m.Value = v),
        new ProjectXmlField<ProjectOutlineCodeLookupValue>("Description", (m, d) => ProjectXmlValue.Text(m.Description), (m, v, d, e) => m.Description = v),
    };
}
