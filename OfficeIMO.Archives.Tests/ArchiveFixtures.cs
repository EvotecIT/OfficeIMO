using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

internal sealed partial class ArchiveFixtures : IDisposable {
    internal ArchiveFixtures() {
        Root = Path.Combine(Path.GetTempPath(), "OfficeIMO.Archives.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Root);
    }
    internal string Root { get; }
    internal string Destination => Path.Combine(Root, "output");
    internal string Package(byte[] bytes) { string path = Path.Combine(Root, "package.bin"); File.WriteAllBytes(path, bytes); return path; }
    public void Dispose() { if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true); }
    internal static byte[] StoredCab(string name, byte[] data) => StoredCab((name, data));
    internal static byte[] Utf8Cabinet(string name) {
        byte[] encodedName = Encoding.UTF8.GetBytes(name + "\0");
        int dataOffset = 44 + 16 + encodedName.Length;
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true);
        writer.Write(Encoding.ASCII.GetBytes("MSCF")); writer.Write(0u); writer.Write((uint)(dataOffset + 9));
        writer.Write(0u); writer.Write(44u); writer.Write(0u);
        writer.Write((byte)3); writer.Write((byte)1); writer.Write((ushort)1); writer.Write((ushort)1);
        writer.Write((ushort)0); writer.Write(0u);
        writer.Write((uint)dataOffset); writer.Write((ushort)1); writer.Write((ushort)0);
        writer.Write(1u); writer.Write(0u); writer.Write((ushort)0); writer.Write(0u); writer.Write((ushort)0x80);
        writer.Write(encodedName);
        writer.Write(0u); writer.Write((ushort)1); writer.Write((ushort)1); writer.Write((byte)'A');
        return stream.ToArray();
    }
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
        bool twoCabinets = false, bool backwardMedia = false, bool lowercaseCabinetKey = false,
        int cabinetCount = 0, bool emptyPayloads = false, bool emptySecondPayload = false, int emptySecondCompression = 0,
        byte[]? unmappedCabinet = null) {
        const string key = "PayloadKey";
        byte[] payload = Encoding.UTF8.GetBytes("hello");
        var schemas = MsiSchemas(wideMedia);
        var rows = new Dictionary<string, object?[][]> {
            ["Directory"] = directoryRows ?? new[] { new object?[] { "TARGETDIR", null, "SourceDir" }, new object?[] { "POLICY", directoryCycle ? "LOCALE" : "TARGETDIR", "POLICY~1|PolicyDefinitions" }, new object?[] { "LOCALE", "POLICY", "EN-US|en-US" } },
            ["Component"] = new[] { new object?[] { "COMP", null, firstDirectory, 0, null, key } },
            ["File"] = new[] { new object?[] { key, "COMP", "TEMPLA~1.ADM|Template.adml", wrongSize ? 6 : 5, null, null, 16384, 1 } },
            ["Media"] = new[] { new object?[] { 1, 1, null, externalCab ? "external.cab" : "#payload.cab", null, null } }
        };
        if (twoCabinets) {
            rows["Component"] = rows["Component"].Concat(new[] { new object?[] { "COMP2", null, secondDirectory, 0, null, "PayloadTwo" } }).ToArray();
            rows["File"] = rows["File"].Concat(new[] { new object?[] { "PayloadTwo", "COMP2", "Second.adml", emptySecondPayload ? 0 : 5, null, null, 16384, 2 } }).ToArray();
            rows["Media"] = backwardMedia
                ? new[] { new object?[] { 1, 2, null, "#second.cab", null, null }, new object?[] { 2, 1, null, "#payload.cab", null, null } }
                : new[] { new object?[] { 2, 2, null, "#second.cab", null, null }, new object?[] { 1, 1, null, "#payload.cab", null, null } };
        }
        if (cabinetCount > 0) {
            rows["Component"] = Enumerable.Range(1, cabinetCount).Select(index => new object?[] {
                "COMP" + index, null, firstDirectory, 0, null, "Payload" + index }).ToArray();
            rows["File"] = Enumerable.Range(1, cabinetCount).Select(index => new object?[] {
                "Payload" + index, "COMP" + index, "Template" + index + ".adml", emptyPayloads ? 0 : 5,
                null, null, 16384, index }).ToArray();
            rows["Media"] = Enumerable.Range(1, cabinetCount).Select(index => new object?[] {
                index, index, null, "#payload" + index + ".cab", null, null }).ToArray();
        }
        var streams = new List<OfficeCompoundStream>();
        if (unmappedCabinet != null) {
            rows["Media"] = rows["Media"].Concat(new[] { new object?[] { 2, 2, null, "#unmapped.cab", null, null } }).ToArray();
            streams.Add(new OfficeCompoundStream(MsiPackageReader.EncodeStreamName("unmapped.cab", false), unmappedCabinet));
        }
        if (cabinetCount > 0) {
            for (int index = 1; index <= cabinetCount; index++) {
                byte[] content = emptyPayloads ? Array.Empty<byte>() : payload;
                // A zero-length file needs no CFDATA block; a cbUncomp=0 record denotes unsupported split data.
                byte[] cabinet = emptyPayloads ? EmptyStoredCab("Payload" + index) : StoredCab("Payload" + index, content);
                streams.Add(new OfficeCompoundStream(MsiPackageReader.EncodeStreamName("payload" + index + ".cab", false), cabinet));
            }
        } else {
            if (!missingCabinet) streams.Add(new OfficeCompoundStream(MsiPackageReader.EncodeStreamName("payload.cab", false), StoredCab(lowercaseCabinetKey ? key.ToLowerInvariant() : key, payload)));
            if (twoCabinets) streams.Add(new OfficeCompoundStream(MsiPackageReader.EncodeStreamName("second.cab", false),
                emptySecondPayload ? EmptyStoredCab("PayloadTwo", emptySecondCompression) : StoredCab("PayloadTwo", Encoding.UTF8.GetBytes("world"))));
        }
        return WriteMsi(schemas, rows, streams, longStrings);
    }
    internal static byte[] EmptyStoredCab(string name, int compression = 0) {
        // With no folder bytes there are no CFDATA records; cbUncomp=0 would denote split data instead.
        byte[] cabinet = StoredCab(name, Array.Empty<byte>());
        Array.Resize(ref cabinet, cabinet.Length - 8);
        for (int index = 0; index < 4; index++) cabinet[8 + index] = (byte)(cabinet.Length >> (index * 8));
        cabinet[40] = 0; cabinet[41] = 0;
        cabinet[42] = (byte)compression; cabinet[43] = (byte)(compression >> 8);
        return cabinet;
    }
}
