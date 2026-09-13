namespace OfficeIMO.Project;

internal static partial class ProjectModelSnapshot {
    private static void Fields(ProjectResourceAvailability item, string path, ProjectModelValues values) {
        values[path + "/From"] = item.From; values[path + "/Through"] = item.Through; values[path + "/Units"] = item.Units;
    }
    private static void Fields(ProjectResourceRate item, string path, ProjectModelValues values) {
        values[path + "/From"] = item.From; values[path + "/To"] = item.To; values[path + "/Table"] = item.Table;
        values[path + "/StandardRate"] = item.StandardRate; values[path + "/OvertimeRate"] = item.OvertimeRate; values[path + "/CostPerUse"] = item.CostPerUse;
        values[path + "/StandardRateFormat"] = item.StandardRateFormat; values[path + "/OvertimeRateFormat"] = item.OvertimeRateFormat;
    }
}
