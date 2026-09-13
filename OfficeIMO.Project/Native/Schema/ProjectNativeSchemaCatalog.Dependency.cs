namespace OfficeIMO.Project;

// MPP14 field descriptors observed in controlled Project 2024 fixtures.
// These are storage-schema definitions, with no project records or source-file payload.
internal static partial class ProjectNativeSchemaCatalog {
    internal static ProjectNativeSchema Dependency() => new ProjectNativeSchema(8, new ProjectNativeSchemaField[] {
        new ProjectNativeSchemaField(0x0E400000u, 0, 0, 10, 32, 3, 4, 0),
        new ProjectNativeSchemaField(0x0E400010u, 1, 65535, 19, 4128, 100, 2, 0),
        new ProjectNativeSchemaField(0x0E400002u, 0, 4, 10, 40, 3, 4, 0),
        new ProjectNativeSchemaField(0x0E400005u, 0, 8, 10, 40, 3, 4, 0),
        new ProjectNativeSchemaField(0x0E400007u, 0, 12, 10, 40, 2, 2, 0),
        new ProjectNativeSchemaField(0x0E400008u, 2, 65535, 18, 32, 100, 2, 1),
        new ProjectNativeSchemaField(0x0E400009u, 0, 14, 10, 40, 3, 4, 0),
        new ProjectNativeSchemaField(0x0E40000Au, 0, 18, 10, 40, 2, 2, 0),
        new ProjectNativeSchemaField(0x0E400015u, 0, 0, 10, 4128, 72, 16, 0),
        new ProjectNativeSchemaField(0x0E400016u, 0, 16, 10, 40, 72, 16, 0),
        new ProjectNativeSchemaField(0x0E400017u, 0, 32, 10, 40, 72, 16, 0),
        new ProjectNativeSchemaField(0x0E40001Cu, 32, 65535, 18, 32, 100, 2, 5),
        new ProjectNativeSchemaField(0x0E40001Du, 0, 33685503, 4, 4298, 8, 0, 0),
    });
}
