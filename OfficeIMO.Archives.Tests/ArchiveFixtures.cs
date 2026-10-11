using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

internal sealed class ArchiveFixtures : IDisposable {
    internal ArchiveFixtures() {
        Root = Path.Combine(Path.GetTempPath(), "OfficeIMO.Archives.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Root);
    }
    internal string Root { get; }
    internal string Destination => Path.Combine(Root, "output");
    internal string Package(byte[] bytes) { string path = Path.Combine(Root, "package.bin"); File.WriteAllBytes(path, bytes); return path; }
    public void Dispose() { if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true); }
    internal static byte[] StoredCab(string name, byte[] data) => StoredCab((name, data));
    internal static byte[] StoredCab(params (string Name, byte[] Data)[] files) {
        byte[][] names = files.Select(file => Encoding.ASCII.GetBytes(file.Name + "\0")).ToArray();
        int payloadOffset = 44 + 16 * files.Length + names.Sum(name => name.Length);
        byte[] data = files.SelectMany(file => file.Data).ToArray();
        using var memory = new MemoryStream(); using var writer = new BinaryWriter(memory);
        writer.Write(Encoding.ASCII.GetBytes("MSCF")); writer.Write(0); writer.Write(payloadOffset + 8 + data.Length); writer.Write(0);
        writer.Write(44); writer.Write(0); writer.Write((byte)3); writer.Write((byte)1); writer.Write((ushort)1); writer.Write((ushort)files.Length);
        writer.Write((ushort)0); writer.Write((ushort)0); writer.Write((ushort)0);
        writer.Write(payloadOffset); writer.Write((ushort)1); writer.Write((ushort)0);
        int offset = 0;
        for (int index = 0; index < files.Length; index++) {
            writer.Write(files[index].Data.Length); writer.Write(offset); writer.Write((ushort)0); writer.Write((ushort)0); writer.Write((ushort)0); writer.Write((ushort)0);
            writer.Write(names[index]); offset += files[index].Data.Length;
        }
        writer.Write(0); writer.Write((ushort)data.Length); writer.Write((ushort)data.Length); writer.Write(data);
        return memory.ToArray();
    }
    internal static byte[] Msi(bool wideMedia = false, bool longStrings = false, bool externalCab = false,
        bool directoryCycle = false, bool wrongSize = false, bool missingCabinet = false,
        object?[][]? directoryRows = null, string firstDirectory = "LOCALE", string secondDirectory = "LOCALE",
        bool twoCabinets = false, bool backwardMedia = false, bool lowercaseCabinetKey = false) {
        const string key = "PayloadKey";
        byte[] payload = Encoding.UTF8.GetBytes("hello");
        var schemas = new Dictionary<string, (string Name, char Kind)[]> {
            ["Directory"] = new[] { ("Directory", 's'), ("Directory_Parent", 's'), ("DefaultDir", 's') },
            ["Component"] = new[] { ("Component", 's'), ("ComponentId", 's'), ("Directory_", 's'), ("Attributes", 'i'), ("Condition", 's'), ("KeyPath", 's') },
            ["File"] = new[] { ("File", 's'), ("Component_", 's'), ("FileName", 's'), ("FileSize", 'I'), ("Version", 's'), ("Language", 's'), ("Attributes", 'i'), ("Sequence", 'i') },
            ["Media"] = new[] { ("DiskId", 'i'), ("LastSequence", wideMedia ? 'I' : 'i'), ("DiskPrompt", 's'), ("Cabinet", 's'), ("VolumeLabel", 's'), ("Source", 's') }
        };
        var rows = new Dictionary<string, object?[][]> {
            ["Directory"] = directoryRows ?? new[] { new object?[] { "TARGETDIR", null, "SourceDir" }, new object?[] { "POLICY", directoryCycle ? "LOCALE" : "TARGETDIR", "POLICY~1|PolicyDefinitions" }, new object?[] { "LOCALE", "POLICY", "EN-US|en-US" } },
            ["Component"] = new[] { new object?[] { "COMP", null, firstDirectory, 0, null, key } },
            ["File"] = new[] { new object?[] { key, "COMP", "TEMPLA~1.ADM|Template.adml", wrongSize ? 6 : 5, null, null, 16384, 1 } },
            ["Media"] = new[] { new object?[] { 1, 1, null, externalCab ? "external.cab" : "#payload.cab", null, null } }
        };
        if (twoCabinets) {
            rows["Component"] = rows["Component"].Concat(new[] { new object?[] { "COMP2", null, secondDirectory, 0, null, "PayloadTwo" } }).ToArray();
            rows["File"] = rows["File"].Concat(new[] { new object?[] { "PayloadTwo", "COMP2", "Second.adml", 5, null, null, 16384, 2 } }).ToArray();
            rows["Media"] = backwardMedia
                ? new[] { new object?[] { 1, 2, null, "#second.cab", null, null }, new object?[] { 2, 1, null, "#payload.cab", null, null } }
                : new[] { new object?[] { 2, 2, null, "#second.cab", null, null }, new object?[] { 1, 1, null, "#payload.cab", null, null } };
        }
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
        var streams = new List<OfficeCompoundStream>();
        using (var pool = new MemoryStream()) using (var data = new MemoryStream()) {
            using var poolWriter = new BinaryWriter(pool, Encoding.UTF8, leaveOpen: true);
            poolWriter.Write((uint)(65001 | (longStrings ? 0x80000000U : 0)));
            foreach (string value in strings) {
                byte[] bytes = Encoding.UTF8.GetBytes(value);
                poolWriter.Write(bytes.Length > ushort.MaxValue ? (ushort)0 : (ushort)bytes.Length);
                poolWriter.Write((ushort)1);
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
        if (!missingCabinet) streams.Add(new OfficeCompoundStream(MsiPackageReader.EncodeStreamName("payload.cab", false), StoredCab(lowercaseCabinetKey ? key.ToLowerInvariant() : key, payload)));
        if (twoCabinets) streams.Add(new OfficeCompoundStream(MsiPackageReader.EncodeStreamName("second.cab", false), StoredCab("PayloadTwo", Encoding.UTF8.GetBytes("world"))));
        return OfficeCompoundFileWriter.Write(streams, new Guid("000c1084-0000-0000-c000-000000000046"));
    }
}
