namespace OfficeIMO.Project;

// MPP14 field descriptors observed in controlled Project 2024 fixtures.
// These are storage-schema definitions, with no project records or source-file payload.
internal static partial class ProjectNativeSchemaCatalog {
    internal static ProjectNativeSchema Calendar() => new ProjectNativeSchema(9, new ProjectNativeSchemaField[] {
        new ProjectNativeSchemaField(0x0D400000u, 1, 65535, 18, 32, 100, 2, 0),
        new ProjectNativeSchemaField(0x0D400006u, 0, 0, 10, 40, 3, 4, 0),
        new ProjectNativeSchemaField(0x0D400007u, 0, 4, 10, 32, 3, 4, 0),
        new ProjectNativeSchemaField(0x0D400009u, 0, 8, 10, 32, 3, 4, 0),
        new ProjectNativeSchemaField(0x0D40000Bu, 1, 65535, 18, 4128, 100, 2, 0),
        new ProjectNativeSchemaField(0x0D400017u, 2, 65535, 18, 4136, 100, 2, 1),
        new ProjectNativeSchemaField(0x0D400008u, 0, 33685503, 0, 4170, 29, 0, 0),
        new ProjectNativeSchemaField(0x0D400001u, 0, 33751039, 4, 202, 8, 512, 0),
        new ProjectNativeSchemaField(0x0D400013u, 0, 33816575, 6, 4330, 3, 4, 0),
        new ProjectNativeSchemaField(0x0D400019u, 4, 65535, 18, 32, 100, 2, 2),
        new ProjectNativeSchemaField(0x0D40001Bu, 0, 0, 10, 4128, 72, 16, 0),
        new ProjectNativeSchemaField(0x0D40001Cu, 0, 16, 10, 32, 72, 16, 0),
        new ProjectNativeSchemaField(0x0D40001Du, 0, 32, 10, 32, 72, 16, 0),
        new ProjectNativeSchemaField(0x0D400022u, 8, 65535, 18, 32, 100, 2, 3),
        new ProjectNativeSchemaField(0x0D400026u, 0, 33882111, 0, 4170, 29, 0, 0),
        new ProjectNativeSchemaField(0x0D40001Au, 0, 33947647, 6, 234, 3, 4, 0),
        new ProjectNativeSchemaField(0x0D40001Eu, 0, 34013183, 4, 202, 8, 104, 0),
        new ProjectNativeSchemaField(0x0D400023u, 0, 34078719, 6, 232, 19, 4, 0),
        new ProjectNativeSchemaField(0x0D400024u, 0, 34144255, 6, 232, 3, 4, 0),
        new ProjectNativeSchemaField(0x0D400025u, 0, 34209791, 4, 202, 8, 512, 0),
        new ProjectNativeSchemaField(0x0D400027u, 0, 34275327, 4, 202, 8, 0, 0),
    });
}
