using System.IO.Compression;
using System.Text;

namespace OfficeIMO.Provenance.Benchmarks;

internal sealed record ProvenanceBenchmarkFixture(
    string Format,
    string Scale,
    string FileName,
    byte[] Asset,
    int ExpectedOutputBytes);

internal static class ProvenanceBenchmarkCorpus {
    internal static readonly string[] Formats = ["PNG", "TIFF", "SVG", "ZIP", "Text"];
    internal static readonly string[] Scales = ["Small", "Large"];

    internal static ProvenanceBenchmarkFixture Create(string format, string scale) {
        int manifestBytes = scale switch {
            "Small" => 4 * 1024,
            "Large" => 1024 * 1024,
            _ => throw new ArgumentOutOfRangeException(nameof(scale))
        };
        byte[] manifest = CreateManifestStore(manifestBytes);
        return format switch {
            "PNG" => CreatePng(scale, manifest),
            "TIFF" => CreateTiff(scale, manifest),
            "SVG" => CreateSvg(scale, manifest),
            "ZIP" => CreateZip(scale, manifest),
            "Text" => CreateText(scale, manifest),
            _ => throw new ArgumentOutOfRangeException(nameof(format))
        };
    }

    private static ProvenanceBenchmarkFixture CreatePng(string scale, byte[] manifest) {
        byte[] header = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        byte[] ihdr = CreatePngChunk("IHDR", [0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0]);
        byte[] carrier = CreatePngChunk("caBX", manifest);
        byte[] image = CreatePngChunk("IDAT", [0x78, 0x9C, 0x63, 0x60, 0x60, 0x60, 0, 0, 0, 4, 0, 1]);
        byte[] end = CreatePngChunk("IEND", []);
        byte[] asset = Join(header, ihdr, carrier, image, end);
        return new ProvenanceBenchmarkFixture("PNG", scale, "fixture.png", asset, asset.Length - carrier.Length);
    }

    private static ProvenanceBenchmarkFixture CreateTiff(string scale, byte[] manifest) {
        const int payloadOffset = 74;
        int pixelOffset = payloadOffset + manifest.Length;
        byte[] asset = new byte[pixelOffset + 1];
        asset[0] = asset[1] = (byte)'I';
        asset[2] = 42;
        asset[4] = 8;
        asset[8] = 5;
        WriteLittleEndianEntry(asset, 10, 256, 4, 1, 1);
        WriteLittleEndianEntry(asset, 22, 257, 4, 1, 1);
        WriteLittleEndianEntry(asset, 34, 273, 4, 1, pixelOffset);
        WriteLittleEndianEntry(asset, 46, 279, 4, 1, 1);
        WriteLittleEndianEntry(asset, 58, 0xCD41, 7, manifest.Length, payloadOffset);
        manifest.CopyTo(asset, payloadOffset);
        return new ProvenanceBenchmarkFixture("TIFF", scale, "fixture.tiff", asset, asset.Length);
    }

    private static ProvenanceBenchmarkFixture CreateSvg(string scale, byte[] manifest) {
        const string prefix = "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:c2pa=\"http://c2pa.org/manifest\"><metadata><c2pa:manifest>";
        const string suffix = "</c2pa:manifest></metadata><text>preserve</text></svg>";
        byte[] asset = Encoding.UTF8.GetBytes(prefix + Convert.ToBase64String(manifest) + suffix);
        int expected = Encoding.UTF8.GetByteCount("<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:c2pa=\"http://c2pa.org/manifest\"><metadata /><text>preserve</text></svg>");
        return new ProvenanceBenchmarkFixture("SVG", scale, "fixture.svg", asset, expected);
    }

    private static ProvenanceBenchmarkFixture CreateZip(string scale, byte[] manifest) {
        byte[] keep = CreateDeterministicBytes(scale == "Large" ? 1024 * 1024 : 4 * 1024);
        byte[] asset = WriteZip(("META-INF/content_credential.c2pa", manifest), ("payload.bin", keep));
        int expectedOutputBytes = OfficeProvenanceRemover.Remove(asset, "fixture.zip").ToArray().Length;
        return new ProvenanceBenchmarkFixture("ZIP", scale, "fixture.zip", asset, expectedOutputBytes);
    }

    private static ProvenanceBenchmarkFixture CreateText(string scale, byte[] manifest) {
        const string before = "before\n";
        const string after = "after\n";
        string block = "-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(manifest) + "\n" +
            "-----END C2PA MANIFEST-----\n";
        byte[] asset = Encoding.UTF8.GetBytes(before + block + after);
        return new ProvenanceBenchmarkFixture("Text", scale, "fixture.md", asset, Encoding.UTF8.GetByteCount(before + after));
    }

    private static byte[] CreateManifestStore(int length) {
        if (length < 284) throw new ArgumentOutOfRangeException(nameof(length));
        int signaturePayloadLength = length - 283;
        byte[] storeDescription = CreateBox("jumd", Join(
            C2paUuid("c2pa"), [0x03], Encoding.ASCII.GetBytes("c2pa\0")));
        byte[] manifestDescription = CreateBox("jumd", Join(
            C2paUuid("c2ma"), [0x03], Encoding.ASCII.GetBytes("m\0")));
        byte[] assertionStoreDescription = CreateBox("jumd", Join(
            C2paUuid("c2as"), [0x03], Encoding.ASCII.GetBytes("c2pa.assertions\0")));
        byte[] assertionDescription = CreateBox("jumd", Join(
            C2paUuid("c2ac"), [0x03], Encoding.ASCII.GetBytes("c2pa.test\0")));
        byte[] assertionStore = CreateBox("jumb", Join(assertionStoreDescription,
            CreateBox("jumb", Join(assertionDescription, CreateBox("cbor", [0xA0])))));
        byte[] claimDescription = CreateBox("jumd", Join(
            C2paUuid("c2cl"), [0x03], Encoding.ASCII.GetBytes("c2pa.claim\0")));
        byte[] claim = CreateBox("jumb", Join(claimDescription, CreateBox("cbor", [0xA0])));
        byte[] signatureDescription = CreateBox("jumd", Join(
            C2paUuid("c2cs"), [0x03], Encoding.ASCII.GetBytes("c2pa.signature\0")));
        byte[] signaturePayload = new byte[signaturePayloadLength];
        Array.Fill(signaturePayload, (byte)0xA0);
        byte[] signature = CreateBox("jumb", Join(signatureDescription, CreateBox("cbor", signaturePayload)));
        return CreateBox("jumb", Join(
            storeDescription,
            CreateBox("jumb", Join(manifestDescription, assertionStore, claim, signature))));
    }

    private static byte[] C2paUuid(string code) => Join(
        Encoding.ASCII.GetBytes(code),
        [0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71]);

    private static byte[] CreateBox(string type, byte[] payload) {
        byte[] box = new byte[payload.Length + 8];
        WriteBigEndian(box, 0, box.Length);
        Encoding.ASCII.GetBytes(type).CopyTo(box, 4);
        payload.CopyTo(box, 8);
        return box;
    }

    private static byte[] CreatePngChunk(string type, byte[] payload) {
        byte[] chunk = new byte[payload.Length + 12];
        WriteBigEndian(chunk, 0, payload.Length);
        Encoding.ASCII.GetBytes(type).CopyTo(chunk, 4);
        payload.CopyTo(chunk, 8);
        WriteBigEndian(chunk, chunk.Length - 4, unchecked((int)ComputePngCrc(chunk, 4, payload.Length + 4)));
        return chunk;
    }

    private static uint ComputePngCrc(byte[] data, int offset, int count) {
        uint crc = 0xFFFFFFFF;
        for (int index = offset; index < offset + count; index++) {
            crc ^= data[index];
            for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? 0xEDB88320U ^ (crc >> 1) : crc >> 1;
        }
        return crc ^ 0xFFFFFFFF;
    }

    private static void WriteLittleEndianEntry(byte[] data, int offset, ushort tag, ushort type, int count, int valueOffset) {
        BitConverter.GetBytes(tag).CopyTo(data, offset);
        BitConverter.GetBytes(type).CopyTo(data, offset + 2);
        BitConverter.GetBytes(count).CopyTo(data, offset + 4);
        BitConverter.GetBytes(valueOffset).CopyTo(data, offset + 8);
    }

    private static byte[] WriteZip(params (string Name, byte[] Data)[] entries) {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach ((string name, byte[] data) in entries) {
                ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
                using Stream target = entry.Open();
                target.Write(data, 0, data.Length);
            }
        }
        return stream.ToArray();
    }

    private static byte[] CreateDeterministicBytes(int count) {
        byte[] result = new byte[count];
        uint state = 0xC2A0_2026;
        for (int index = 0; index < result.Length; index++) {
            state = unchecked(state * 1664525 + 1013904223);
            result[index] = (byte)(state >> 24);
        }
        return result;
    }

    private static void WriteBigEndian(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private static byte[] Join(params byte[][] parts) {
        int length = 0;
        foreach (byte[] part in parts) length = checked(length + part.Length);
        byte[] result = new byte[length];
        int offset = 0;
        foreach (byte[] part in parts) {
            part.CopyTo(result, offset);
            offset += part.Length;
        }
        return result;
    }
}
