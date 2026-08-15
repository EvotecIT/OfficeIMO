using System.IO.Compression;
using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void StrictXmpRemovalPreservesNonScalarRdfProperties(bool useResource) {
        string value = "http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia";
        string property = useResource
            ? $"<iptc:DigitalSourceType rdf:resource=\"{value}\"><rdf:Description><keep>yes</keep></rdf:Description></iptc:DigitalSourceType>"
            : $"<iptc:DigitalSourceType>{value}<rdf:Description><keep>yes</keep></rdf:Description></iptc:DigitalSourceType>";
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\"><metadata><x:xmpmeta xmlns:x=\"adobe:ns:meta/\">" +
            "<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\"><rdf:Description " +
            "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\">" + property +
            "</rdf:Description></rdf:RDF></x:xmpmeta></metadata></svg>");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.svg");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Contains("<keep>yes</keep>", Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void ZipRewriteUsesValidatedUnicodePathName() {
        byte[] package = CreateZipWithUnicodePathEntry("cafe.txt", "café.txt", Encoding.UTF8.GetBytes("keep"));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "publication.zip");

        Assert.True(result.WasChanged);
        using var archive = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);
        Assert.NotNull(archive.GetEntry("café.txt"));
        Assert.Null(archive.GetEntry("cafe.txt"));
    }

    private static byte[] CreateZipWithUnicodePathEntry(string rawName, string unicodeName, byte[] content) {
        byte[] manifest = CreateManifestStore();
        using var output = new MemoryStream();
        using var writer = new BinaryWriter(output, Encoding.UTF8, leaveOpen: true);
        var records = new List<(byte[] Name, byte[] Extra, byte[] Data, uint Crc, uint Offset)>();
        AddStored(rawName, unicodeName, content);
        AddStored("META-INF/content_credential.c2pa", null, manifest);
        uint centralOffset = checked((uint)output.Position);
        foreach ((byte[] name, byte[] extra, byte[] data, uint crc, uint offset) in records) {
            writer.Write(0x02014B50U); writer.Write((ushort)20); writer.Write((ushort)20); writer.Write((ushort)0);
            writer.Write((ushort)0); writer.Write((ushort)0); writer.Write((ushort)0); writer.Write(crc);
            writer.Write((uint)data.Length); writer.Write((uint)data.Length); writer.Write((ushort)name.Length);
            writer.Write((ushort)extra.Length); writer.Write((ushort)0); writer.Write((ushort)0); writer.Write((ushort)0);
            writer.Write(0U); writer.Write(offset); writer.Write(name); writer.Write(extra);
        }
        uint centralSize = checked((uint)output.Position - centralOffset);
        writer.Write(0x06054B50U); writer.Write((ushort)0); writer.Write((ushort)0); writer.Write((ushort)records.Count);
        writer.Write((ushort)records.Count); writer.Write(centralSize); writer.Write(centralOffset); writer.Write((ushort)0);
        writer.Flush();
        return output.ToArray();

        void AddStored(string name, string? unicode, byte[] data) {
            byte[] raw = Encoding.ASCII.GetBytes(name);
            byte[] extra = unicode == null ? Array.Empty<byte>() : CreateUnicodePathExtra(raw, unicode);
            uint crc = ComputePngCrc(data, 0, data.Length);
            uint offset = checked((uint)output.Position);
            writer.Write(0x04034B50U); writer.Write((ushort)20); writer.Write((ushort)0); writer.Write((ushort)0);
            writer.Write((ushort)0); writer.Write((ushort)0); writer.Write(crc); writer.Write((uint)data.Length);
            writer.Write((uint)data.Length); writer.Write((ushort)raw.Length); writer.Write((ushort)extra.Length);
            writer.Write(raw); writer.Write(extra); writer.Write(data);
            records.Add((raw, extra, data, crc, offset));
        }
    }

    private static byte[] CreateUnicodePathExtra(byte[] rawName, string unicodeName) {
        byte[] utf8 = Encoding.UTF8.GetBytes(unicodeName);
        using var output = new MemoryStream();
        using var writer = new BinaryWriter(output, Encoding.UTF8, leaveOpen: true);
        writer.Write((ushort)0x7075); writer.Write((ushort)(5 + utf8.Length)); writer.Write((byte)1);
        writer.Write(ComputePngCrc(rawName, 0, rawName.Length)); writer.Write(utf8); writer.Flush();
        return output.ToArray();
    }
}
