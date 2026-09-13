namespace OfficeIMO.Project;

// Field storage facts only; no producer records, templates, strings, or process pointers.
internal static partial class ProjectNativeSchema8Catalog {
    internal static ProjectNativeLegacy8Schema Dependency() {
        var schema = new ProjectNativeLegacy8Schema(36, 28, 32, -1);
        schema.Field(0x0e400000u, 0, 32, 3, 4, 0);
        schema.Field(0x0e400010u, 4, 32, 100, 4, 0);
        schema.Field(0x0e400008u, 8, 32, 100, 4, 0);
        schema.Field(0x0e400002u, 12, 40, 3, 4, 0);
        schema.Field(0x0e400005u, 16, 40, 3, 4, 0);
        schema.Field(0x0e400007u, 20, 40, 2, 2, 0);
        schema.Field(0x0e40000au, 22, 40, 2, 2, 0);
        schema.Field(0x0e400009u, 24, 40, 3, 4, 0);
        return schema;
    }
}
