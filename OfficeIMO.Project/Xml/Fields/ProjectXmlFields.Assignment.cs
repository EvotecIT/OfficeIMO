namespace OfficeIMO.Project;

internal static partial class ProjectXmlFields {
    internal static readonly ProjectXmlField<ProjectAssignment>[] Assignment = {
        new ProjectXmlField<ProjectAssignment>("Stop", (m,d) => ProjectXmlValue.Date(m.Stop), (m,v,d,e) => m.Stop = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectAssignment>("Resume", (m,d) => ProjectXmlValue.Date(m.Resume), (m,v,d,e) => m.Resume = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectAssignment>("GUID", (m, d) => ProjectXmlValue.Identifier(m.Guid), (m, v, d, e) => m.Guid = ProjectXmlValue.ParseGuid(v)),
        new ProjectXmlField<ProjectAssignment>("Units", (m, d) => ProjectXmlValue.Units(m.Units), (m, v, d, e) => m.Units = ProjectXmlValue.ParseUnits(v)),
        new ProjectXmlField<ProjectAssignment>("CostRateTable", (m,d) => ProjectXmlValue.Integer((int?)m.CostRateTable), (m,v,d,e) => m.CostRateTable = (ProjectCostRateTable)ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectAssignment>("Delay", (m,d) => ProjectXmlValue.Number(m.DelayMinutes * 10m), (m,v,d,e) => m.DelayMinutes = ProjectXmlValue.ParseNumber(v) / 10m),
        new ProjectXmlField<ProjectAssignment>("HasFixedRateUnits", (m,d) => ProjectXmlValue.Boolean(m.HasFixedRateUnits), (m,v,d,e) => m.HasFixedRateUnits = ProjectXmlValue.ParseBool(v)),
        new ProjectXmlField<ProjectAssignment>("RateScale", (m,d) => ProjectXmlValue.Integer(m.MaterialRateScale), (m,v,d,e) => m.MaterialRateScale = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectAssignment>("WorkContour", (m,d) => ProjectXmlValue.Integer((int?)m.WorkContour), (m,v,d,e) => m.WorkContour = (ProjectWorkContour)ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectAssignment>("Work", (m, d) => ProjectXmlValue.Work(m.Work), (m, v, d, e) => m.Work = ProjectXmlValue.ParseWork(v)),
        new ProjectXmlField<ProjectAssignment>("ActualWork", (m, d) => ProjectXmlValue.Work(m.ActualWork), (m, v, d, e) => m.ActualWork = ProjectXmlValue.ParseWork(v)),
        new ProjectXmlField<ProjectAssignment>("RemainingWork", (m, d) => ProjectXmlValue.Work(m.RemainingWork), (m, v, d, e) => m.RemainingWork = ProjectXmlValue.ParseWork(v)),
        new ProjectXmlField<ProjectAssignment>("OvertimeWork", (m, d) => ProjectXmlValue.Work(m.OvertimeWork), (m, v, d, e) => m.OvertimeWork = ProjectXmlValue.ParseWork(v)),
        new ProjectXmlField<ProjectAssignment>("ActualOvertimeWork", (m, d) => ProjectXmlValue.Work(m.ActualOvertimeWork), (m, v, d, e) => m.ActualOvertimeWork = ProjectXmlValue.ParseWork(v)),
        new ProjectXmlField<ProjectAssignment>("Start", (m, d) => ProjectXmlValue.Date(m.Start), (m, v, d, e) => m.Start = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectAssignment>("Finish", (m, d) => ProjectXmlValue.Date(m.Finish), (m, v, d, e) => m.Finish = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectAssignment>("ActualStart", (m, d) => ProjectXmlValue.Date(m.ActualStart), (m, v, d, e) => m.ActualStart = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectAssignment>("ActualFinish", (m, d) => ProjectXmlValue.Date(m.ActualFinish), (m, v, d, e) => m.ActualFinish = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectAssignment>("Cost", (m, d) => ProjectXmlValue.Money(m.Cost), (m, v, d, e) => m.Cost = ProjectXmlValue.ParseMoney(v)),
        new ProjectXmlField<ProjectAssignment>("ActualCost", (m, d) => ProjectXmlValue.Money(m.ActualCost), (m, v, d, e) => m.ActualCost = ProjectXmlValue.ParseMoney(v)),
        new ProjectXmlField<ProjectAssignment>("RemainingCost", (m, d) => ProjectXmlValue.Money(m.RemainingCost), (m, v, d, e) => m.RemainingCost = ProjectXmlValue.ParseMoney(v)),
        new ProjectXmlField<ProjectAssignment>("PercentWorkComplete", (m, d) => ProjectXmlValue.Integer(m.PercentWorkComplete), (m, v, d, e) => m.PercentWorkComplete = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectAssignment>("Notes", (m, d) => ProjectXmlValue.Text(m.Notes), (m, v, d, e) => m.Notes = v),
    };
}
