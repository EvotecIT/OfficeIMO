using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

internal sealed partial class ArchiveFixtures {
    private static Dictionary<string, (string Name, char Kind)[]> MsiSchemas(bool wideMedia = false) => new() {
        ["Directory"] = new[] { ("Directory", 's'), ("Directory_Parent", 's'), ("DefaultDir", 's') },
        ["Component"] = new[] { ("Component", 's'), ("ComponentId", 's'), ("Directory_", 's'), ("Attributes", 'i'), ("Condition", 's'), ("KeyPath", 's') },
        ["File"] = new[] { ("File", 's'), ("Component_", 's'), ("FileName", 's'), ("FileSize", 'I'), ("Version", 's'), ("Language", 's'), ("Attributes", 'i'), ("Sequence", 'i') },
        ["Media"] = new[] { ("DiskId", 'i'), ("LastSequence", wideMedia ? 'I' : 'i'), ("DiskPrompt", 's'), ("Cabinet", 's'), ("VolumeLabel", 's'), ("Source", 's') }
    };

    private static byte[] WriteMsi(Dictionary<string, (string Name, char Kind)[]> schemas,
        Dictionary<string, object?[][]> rows, List<OfficeCompoundStream> streams, bool longStrings = false) {
        var columns = new List<object?[]>();
        foreach (var table in schemas) for (int column = 0; column < table.Value.Length; column++) {
            var field = table.Value[column];
            columns.Add(new object?[] { table.Key, column + 1, field.Name, field.Kind == 's' ? 0x0d48 : field.Kind == 'I' ? 0x104 : 0x102 });
        }
        schemas.Add("_Columns", new[] { ("Table", 's'), ("Number", 'i'), ("Name", 's'), ("Type", 'i') });
        rows.Add("_Columns", columns.ToArray());
        string[] strings = rows.Values.SelectMany(table => table).SelectMany(row => row).OfType<string>().Distinct().ToArray();
        if (longStrings) strings = Enumerable.Range(0, 65536).Select(index => "unused-" + index)
            .Concat(new[] { new string('x', 70000) }).Concat(strings).ToArray();
        var ids = strings.Select((value, index) => new { value, index }).ToDictionary(item => item.value, item => item.index + 1);
        using (var pool = new MemoryStream()) using (var data = new MemoryStream()) {
            using var poolWriter = new BinaryWriter(pool, Encoding.UTF8, leaveOpen: true);
            poolWriter.Write((uint)(65001 | (longStrings ? 0x80000000U : 0)));
            foreach (string value in strings) {
                byte[] bytes = Encoding.UTF8.GetBytes(value);
                poolWriter.Write(bytes.Length > ushort.MaxValue ? (ushort)0 : (ushort)bytes.Length);
                // Native Windows Installer stores [0, references] followed by the UInt32 length.
                // Keep the 70,000-byte fixture's references (2) unequal to its length high word (1).
                poolWriter.Write(bytes.Length > ushort.MaxValue ? (ushort)2 : (ushort)1);
                if (bytes.Length > ushort.MaxValue) poolWriter.Write((uint)bytes.Length);
                data.Write(bytes, 0, bytes.Length);
            }
            streams.Add(new OfficeCompoundStream(MsiPackageReader.EncodeStreamName("_StringPool", true), pool.ToArray()));
            streams.Add(new OfficeCompoundStream(MsiPackageReader.EncodeStreamName("_StringData", true), data.ToArray()));
        }
        foreach (var table in rows) {
            using var encoded = new MemoryStream(); using var writer = new BinaryWriter(encoded);
            for (int column = 0; column < schemas[table.Key].Length; column++) foreach (object?[] row in table.Value) {
                object? value = row[column]; char kind = schemas[table.Key][column].Kind;
                if (kind == 's') {
                    int id = value == null ? 0 : ids[(string)value]; writer.Write((ushort)id); if (longStrings) writer.Write((byte)(id >> 16));
                } else if (kind == 'I') writer.Write(value == null ? 0U : (uint)((int)value ^ int.MinValue));
                else writer.Write(value == null ? (ushort)0 : (ushort)((int)value ^ 0x8000));
            }
            streams.Add(new OfficeCompoundStream(MsiPackageReader.EncodeStreamName(table.Key, true), encoded.ToArray()));
        }
        return OfficeCompoundFileWriter.Write(streams, new Guid("000c1084-0000-0000-c000-000000000046"));
    }

    internal static byte[] MsiAliasFiles(bool reverseRows, params MsiFixtureFile[] files) => MsiAliasFiles(reverseRows, null, files);

    internal static byte[] MsiAliasFiles(bool reverseRows, int? wordCount, params MsiFixtureFile[] files) =>
        MsiFiles(reverseRows, false, 1, wordCount, files);

    internal static byte[] MsiFiles(bool reverseRows, bool reverseCabinet, int firstDisk, int? wordCount, params MsiFixtureFile[] files) {
        var directoryRows = new List<object?[]> { new object?[] { "TARGETDIR", null, "SourceDir" } };
        var directories = new Dictionary<string, string>(StringComparer.Ordinal) { [string.Empty] = "TARGETDIR" };
        var components = new List<object?[]>(); var fileRows = new List<object?[]>();
        var cabinetFiles = new Dictionary<int, List<(string Name, byte[] Data)>>();
        var media = new List<object?[]>();
        int sequence = 0;
        foreach (var item in files.Select((file, index) => (File: file, Index: index)).OrderBy(item => item.File.Cabinet)) {
            MsiFixtureFile file = item.File;
            string parent = string.Empty;
            foreach (string segment in file.Parent.Split('/')) {
                string path = parent.Length == 0 ? segment : parent + "/" + segment;
                if (!directories.ContainsKey(path)) {
                    string id = "Directory" + directories.Count;
                    directoryRows.Add(new object?[] { id, directories[parent], segment }); directories.Add(path, id);
                }
                parent = path;
            }
            string key = "File" + item.Index; string component = "Component" + item.Index;
            components.Add(new object?[] { component, null, directories[parent], 0, null, key });
            fileRows.Add(new object?[] { key, component, file.Name, file.DeclaredSize ?? file.Data.Length, null, null, file.Attributes, ++sequence });
            if (!cabinetFiles.TryGetValue(file.Cabinet, out var payloads)) cabinetFiles.Add(file.Cabinet, payloads = new());
            payloads.Add((key, file.Data));
        }
        int last = 0; var streams = new List<OfficeCompoundStream>();
        foreach (var cabinet in cabinetFiles.OrderBy(pair => pair.Key)) {
            last += cabinet.Value.Count;
            string name = "payload" + cabinet.Key + ".cab";
            media.Add(new object?[] { cabinet.Key + firstDisk - 1, last, null, "#" + name, null, null });
            var payloads = reverseCabinet ? cabinet.Value.AsEnumerable().Reverse().ToArray() : cabinet.Value.ToArray();
            streams.Add(new OfficeCompoundStream(MsiPackageReader.EncodeStreamName(name, false), StoredCab(payloads)));
        }
        if (reverseRows) fileRows.Reverse();
        var rows = new Dictionary<string, object?[][]> {
            ["Directory"] = directoryRows.ToArray(), ["Component"] = components.ToArray(),
            ["File"] = fileRows.ToArray(), ["Media"] = media.ToArray()
        };
        if (wordCount.HasValue) streams.Add(new OfficeCompoundStream(OfficeOlePropertySetWriter.SummaryInformationStreamName,
            OfficeOlePropertySetWriter.CreatePropertySet((OfficeOlePropertySetWriter.SummaryInformationFormatId,
                OfficeOlePropertySetWriter.CreateSection(new[] { OfficeOleProperty.Int32(15, wordCount.Value) })))));
        return WriteMsi(MsiSchemas(), rows, streams);
    }
}

internal sealed class MsiFixtureFile {
    internal MsiFixtureFile(string name, string text = "hello", int cabinet = 1,
        string parent = "PolicyDefinitions/en-US", int? declaredSize = null, int attributes = 16384) {
        Name = name; Data = Encoding.UTF8.GetBytes(text); Cabinet = cabinet; Parent = parent; DeclaredSize = declaredSize; Attributes = attributes;
    }
    internal string Name { get; }
    internal byte[] Data { get; }
    internal int Cabinet { get; }
    internal string Parent { get; }
    internal int? DeclaredSize { get; }
    internal int Attributes { get; }
}
