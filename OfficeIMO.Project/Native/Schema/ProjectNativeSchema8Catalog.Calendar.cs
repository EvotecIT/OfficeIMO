namespace OfficeIMO.Project;

// Field storage facts only; no producer records, templates, strings, or process pointers.
internal static partial class ProjectNativeSchema8Catalog {
    internal static ProjectNativeLegacy8Schema Calendar() {
        var schema = new ProjectNativeLegacy8Schema(36, 24, 28, 32);
        schema.Field(0x0d400009u, 0, 32, 3, 4, 0);
        schema.Field(0x0d400006u, 4, 40, 3, 4, 0);
        schema.Field(0x0d400007u, 8, 32, 3, 4, 0);
        schema.Field(0x0d400000u, 12, 32, 100, 4, 0);
        schema.Field(0x0d400003u, 12, 40, 100, 4, 1);
        schema.Field(0x0d40000bu, 16, 32, 100, 4, 0);
        schema.Field(0x0d400001u, 20, 202, 8, 512, 0);
        schema.Field(0x0d400013u, 7, 234, 3, 4, 0);
        schema.Field(0x0d400008u, 8, 202, 29, 0, 0);
        return schema;
    }
}
