namespace OfficeIMO.Project;

internal static partial class ProjectXmlFields {
    internal static readonly ProjectXmlField<ProjectResourceAvailability>[] ResourceAvailability = {
        new ProjectXmlField<ProjectResourceAvailability>("AvailableFrom", (m,d) => ProjectXmlValue.Date(m.From), (m,v,d,e) => m.From = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectResourceAvailability>("AvailableTo", (m,d) => ProjectXmlValue.Date(m.Through), (m,v,d,e) => m.Through = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectResourceAvailability>("AvailableUnits", (m,d) => ProjectXmlValue.Units(m.Units), (m,v,d,e) => m.Units = ProjectXmlValue.ParseUnits(v))
    };
    internal static readonly ProjectXmlField<ProjectResourceRate>[] ResourceRate = {
        new ProjectXmlField<ProjectResourceRate>("RatesFrom", (m,d) => ProjectXmlValue.Date(m.From), (m,v,d,e) => m.From = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectResourceRate>("RatesTo", (m,d) => ProjectXmlValue.Date(m.To), (m,v,d,e) => m.To = ProjectXmlValue.ParseDate(v)),
        new ProjectXmlField<ProjectResourceRate>("RateTable", (m,d) => ProjectXmlValue.Integer((int?)m.Table), (m,v,d,e) => m.Table = (ProjectCostRateTable)ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectResourceRate>("StandardRate", (m,d) => ProjectXmlValue.Number(m.StandardRate), (m,v,d,e) => m.StandardRate = ProjectXmlValue.ParseNumber(v)),
        new ProjectXmlField<ProjectResourceRate>("StandardRateFormat", (m,d) => ProjectXmlValue.Integer(m.StandardRateFormat), (m,v,d,e) => m.StandardRateFormat = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectResourceRate>("OvertimeRate", (m,d) => ProjectXmlValue.Number(m.OvertimeRate), (m,v,d,e) => m.OvertimeRate = ProjectXmlValue.ParseNumber(v)),
        new ProjectXmlField<ProjectResourceRate>("OvertimeRateFormat", (m,d) => ProjectXmlValue.Integer(m.OvertimeRateFormat), (m,v,d,e) => m.OvertimeRateFormat = ProjectXmlValue.ParseInt(v)),
        new ProjectXmlField<ProjectResourceRate>("CostPerUse", (m,d) => ProjectXmlValue.Money(m.CostPerUse), (m,v,d,e) => m.CostPerUse = ProjectXmlValue.ParseMoney(v))
    };
}
