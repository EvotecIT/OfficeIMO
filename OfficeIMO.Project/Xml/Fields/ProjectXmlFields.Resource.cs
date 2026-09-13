namespace OfficeIMO.Project;

internal static partial class ProjectXmlFields {
    internal static readonly ProjectXmlField<ProjectResource>[] Resource = {
        new ProjectXmlField<ProjectResource>("GUID", (m, d) => ProjectXmlValue.Identifier(m.Guid), (m, v, d, e) => m.Guid = ProjectXmlValue.ParseGuid(v)),
        new ProjectXmlField<ProjectResource>("Name", (m, d) => ProjectXmlValue.Text(m.Name), (m, v, d, e) => m.Name = v),
        new ProjectXmlField<ProjectResource>("ID", (m, d) => ProjectXmlValue.Integer(m.DisplayId), (m, v, d, e) => m.DisplayId = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectResource>("Type", (m, d) => m.Type.HasValue ? ProjectXmlValue.Integer(m.Type == ProjectResourceType.Cost ? 0 : (int)m.Type.Value) : null,
            (m, v, d, e) => m.Type = (bool?)e.Element(e.Name.Namespace + "IsCostResource") == true ? ProjectResourceType.Cost : (ProjectResourceType)ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectResource>("IsCostResource", (m, d) => m.Type.HasValue ? ProjectXmlValue.Boolean(m.Type == ProjectResourceType.Cost) : null,
            (m, v, d, e) => { if (ProjectXmlValue.ParseBool(v)) m.Type = ProjectResourceType.Cost; }),
        new ProjectXmlField<ProjectResource>("Initials", (m, d) => ProjectXmlValue.Text(m.Initials), (m, v, d, e) => m.Initials = v),
        new ProjectXmlField<ProjectResource>("Group", (m, d) => ProjectXmlValue.Text(m.Group), (m, v, d, e) => m.Group = v),
        new ProjectXmlField<ProjectResource>("EmailAddress", (m, d) => ProjectXmlValue.Text(m.EmailAddress), (m, v, d, e) => m.EmailAddress = v),
        new ProjectXmlField<ProjectResource>("MaterialLabel", (m, d) => ProjectXmlValue.Text(m.MaterialLabel), (m, v, d, e) => m.MaterialLabel = v),
        new ProjectXmlField<ProjectResource>("MaxUnits", (m, d) => ProjectXmlValue.Units(m.MaxUnits), (m, v, d, e) => m.MaxUnits = ProjectXmlValue.ParseUnits(v)),
        new ProjectXmlField<ProjectResource>("AccrueAt", (m,d) => ProjectXmlValue.Integer((int?)m.AccrueAt), (m,v,d,e) => m.AccrueAt = (ProjectCostAccrual)ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectResource>("CanLevel", (m,d) => ProjectXmlValue.Boolean(m.CanLevel), (m,v,d,e) => m.CanLevel = ProjectXmlValue.ParseBool(v)),
        new ProjectXmlField<ProjectResource>("StandardRate", (m, d) => ProjectXmlValue.Number(m.StandardRate), (m, v, d, e) => m.StandardRate = ProjectXmlValue.ParseNumber(v)),
        new ProjectXmlField<ProjectResource>("OvertimeRate", (m, d) => ProjectXmlValue.Number(m.OvertimeRate), (m, v, d, e) => m.OvertimeRate = ProjectXmlValue.ParseNumber(v)),
        new ProjectXmlField<ProjectResource>("CostPerUse", (m, d) => ProjectXmlValue.Money(m.CostPerUse), (m, v, d, e) => m.CostPerUse = ProjectXmlValue.ParseMoney(v)),
        new ProjectXmlField<ProjectResource>("Cost", (m, d) => ProjectXmlValue.Money(m.Cost), (m, v, d, e) => m.Cost = ProjectXmlValue.ParseMoney(v)),
        new ProjectXmlField<ProjectResource>("ActualCost", (m, d) => ProjectXmlValue.Money(m.ActualCost), (m, v, d, e) => m.ActualCost = ProjectXmlValue.ParseMoney(v)),
        new ProjectXmlField<ProjectResource>("RemainingCost", (m, d) => ProjectXmlValue.Money(m.RemainingCost), (m, v, d, e) => m.RemainingCost = ProjectXmlValue.ParseMoney(v)),
        new ProjectXmlField<ProjectResource>("Work", (m, d) => ProjectXmlValue.Work(m.Work), (m, v, d, e) => m.Work = ProjectXmlValue.ParseWork(v)),
        new ProjectXmlField<ProjectResource>("ActualWork", (m, d) => ProjectXmlValue.Work(m.ActualWork), (m, v, d, e) => m.ActualWork = ProjectXmlValue.ParseWork(v)),
        new ProjectXmlField<ProjectResource>("RemainingWork", (m, d) => ProjectXmlValue.Work(m.RemainingWork), (m, v, d, e) => m.RemainingWork = ProjectXmlValue.ParseWork(v)),
        new ProjectXmlField<ProjectResource>("Notes", (m, d) => ProjectXmlValue.Text(m.Notes), (m, v, d, e) => m.Notes = v),
        new ProjectXmlField<ProjectResource>("IsNull", (m, d) => ProjectXmlValue.Boolean(m.IsNull), (m, v, d, e) => m.IsNull = ProjectXmlValue.ParseBool(v)),
    };
}
