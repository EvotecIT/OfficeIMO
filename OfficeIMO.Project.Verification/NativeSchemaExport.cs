using OfficeIMO.Core.Internal;
using OfficeIMO.Project;
using System.Globalization;
using System.Text;

/// <summary>Exports field definitions from controlled producer fixtures, never document records.</summary>
internal static class NativeSchemaExport {
    internal static int Run(string input, string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        if (!OfficeCompoundFileReader.TryRead(File.ReadAllBytes(input), new OfficeCompoundReadOptions(4096, 2048, 32 * 1024 * 1024, 64 * 1024 * 1024),
            out var file, out var error) || file == null) throw new InvalidDataException(error);
        var profile = ProjectNativeProfile.Detect(file);
        var properties = ProjectNativeProperties.Read(file.Streams[profile.Properties], default, profile == ProjectNativeProfile.Mpp8);
        Directory.CreateDirectory(output);
        if (profile == ProjectNativeProfile.Mpp8) return Legacy8(properties, output);
        string catalog = "ProjectNativeSchema" + profile.Version + "Catalog";
        foreach (var table in new[] { ("Task", 0x14u), ("Resource", 0x15u), ("Calendar", 0x16u), ("Assignment", 0x17u), ("Dependency", 0x18u), ("OutlineCode", 0x19u) }) {
            var primary = properties[0x03000000 | table.Item2];
            var extended = properties.TryGetValue(0x00020000 | table.Item2, out var second) ? second : primary;
            if (primary.Length % 28 != 0 || extended.Length % 28 != 0 || extended.Length < primary.Length ||
                !extended.Slice(0, primary.Length).Copy().SequenceEqual(primary.Copy())) throw new InvalidDataException("Inconsistent producer field maps.");
            var source = new StringBuilder("namespace OfficeIMO.Project;\n\n// Generated from producer field metadata by native-schema.\n// Storage definitions only; no document records or template payload.\ninternal static partial class " + catalog + " {\n");
            source.Append("    internal static ProjectNativeSchema ").Append(table.Item1).Append("() => new ProjectNativeSchema(")
                .Append(primary.Length / 28).Append(", new ProjectNativeSchemaField[] {\n");
            for (int i = 0; i < extended.Length; i += 28) {
                source.Append("        new ProjectNativeSchemaField(0x").Append(extended.UInt32(i + 12).ToString("X8", CultureInfo.InvariantCulture)).Append("u, ");
                source.Append(extended.UInt32(i)).Append(", ").Append(extended.Int32(i + 4)).Append(", ").Append(extended.Int32(i + 8)).Append(", ")
                    .Append(extended.Int32(i + 16)).Append(", ").Append(extended.UInt16(i + 20)).Append(", ").Append(extended.UInt16(i + 22)).Append(", ")
                    .Append(extended.Int32(i + 24)).Append("),\n");
            }
            source.Append("    });\n}\n");
            File.WriteAllText(Path.Combine(output, catalog + "." + table.Item1 + ".cs"), source.ToString());
        }
        return 0;
    }

    private static int Legacy8(Dictionary<uint, ProjectNativeValue> properties, string output) {
        uint ordinal = 0;
        foreach (string name in new[] { "Task", "Resource", "Calendar", "Assignment", "Dependency" }) {
            var layout = ProjectNativeProperties.Read(properties[0x02000000u | ++ordinal].Copy(), default);
            var fields = layout[1];
            if (fields.Length % 24 != 0) throw new InvalidDataException("Unrecognized Project 98 descriptor width.");
            var source = new StringBuilder("namespace OfficeIMO.Project;\n\n// Field storage facts only; no producer records, templates, strings, or process pointers.\ninternal static partial class ProjectNativeSchema8Catalog {\n");
            source.Append("    internal static ProjectNativeLegacy8Schema ").Append(name).Append("() {\n        var schema = new ProjectNativeLegacy8Schema(")
                .Append(layout[5].Int32()).Append(", ").Append(layout[6].Int32()).Append(", ").Append(layout[7].Int32()).Append(", ")
                .Append(layout[8].Int32()).Append(");\n");
            for (int offset = 0; offset < fields.Length; offset += 24)
                source.Append("        schema.Field(0x").Append(fields.UInt32(offset + 8).ToString("x8", CultureInfo.InvariantCulture)).Append("u, ")
                    .Append(fields.UInt16(offset + 4)).Append(", ").Append(fields.Int32(offset + 12)).Append(", ").Append(fields.UInt16(offset + 16))
                    .Append(", ").Append(fields.UInt16(offset + 18)).Append(", ").Append(fields.UInt16(offset + 20)).Append(");\n");
            source.Append("        return schema;\n    }\n}\n");
            File.WriteAllText(Path.Combine(output, "ProjectNativeSchema8Catalog." + name + ".cs"), source.ToString());
        }
        return 0;
    }
}
