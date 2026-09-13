namespace OfficeIMO.Project;

internal static partial class ProjectModelSnapshot {
    private static void Fields(ProjectOutlineCodeDefinition item, string path, ProjectModelValues values) {
        values[path + "/Guid"] = item.Guid;
        values[path + "/FieldId"] = item.FieldId;
        values[path + "/FieldName"] = item.FieldName;
        values[path + "/Alias"] = item.Alias;
        values[path + "/LeafOnly"] = item.LeafOnly;
        values[path + "/AllLevelsRequired"] = item.AllLevelsRequired;
        values[path + "/OnlyTableValuesAllowed"] = item.OnlyTableValuesAllowed;
    }
    private static void Fields(ProjectOutlineCodeMask item, string path, ProjectModelValues values) {
        values[path + "/Level"] = item.Level;
        values[path + "/Type"] = item.Type;
        values[path + "/Length"] = item.Length;
        values[path + "/Separator"] = item.Separator;
    }
    private static void Fields(ProjectOutlineCodeLookupValue item, string path, ProjectModelValues values) {
        values[path + "/ValueId"] = item.ValueId;
        values[path + "/Guid"] = item.Guid;
        values[path + "/ParentValueId"] = item.ParentValueId;
        values[path + "/Type"] = item.Type;
        values[path + "/Value"] = item.Value;
        values[path + "/Description"] = item.Description;
    }
}
