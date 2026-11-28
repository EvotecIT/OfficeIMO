using System.IO.Packaging;
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointEmbeddedPayloadBenchmarks {
    private const int SlideCount = 12;
    private const int PayloadLength = 128 * 1024;
    private const string PayloadRelationshipId = "rIdAttachedDocument";
    private const string FirstPayloadIdentity = "/ppt/slides/slide1.xml#" + PayloadRelationshipId;
    private const string PackageRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
    private byte[] _deck = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? path = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_PAYLOAD_FIXTURE");
        _deck = string.IsNullOrEmpty(path) ? CreateFixture() : File.ReadAllBytes(path);
        IReadOnlyList<OfficeEmbeddedPayloadInfo> manifest = ReadPayloadInventory();
        IReadOnlyList<OfficeEmbeddedPayloadInfo> hashes = HashEmbeddedPayloads();
        if (manifest.Count != SlideCount || manifest.Any(info => info.Kind != OfficeEmbeddedPayloadKind.EmbeddedPackage || info.HasKnownLength) ||
            hashes.Count != SlideCount || hashes.Any(info => info.Length != PayloadLength || info.Sha256?.Length != 64) ||
            ReadDetachedPayload() != PayloadLength) {
            throw new InvalidOperationException("The benchmark fixture does not contain the expected embedded documents.");
        }
    }

    [Benchmark]
    public IReadOnlyList<OfficeEmbeddedPayloadInfo> ReadPayloadInventory() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        using OfficeEmbeddedPayloadInspection inspection = document.InspectEmbeddedPayloads();
        return inspection.Payloads;
    }

    [Benchmark]
    public IReadOnlyList<OfficeEmbeddedPayloadInfo> HashEmbeddedPayloads() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        return document.GetEmbeddedPayloads(includeSha256: true);
    }

    [Benchmark]
    public long ReadDetachedPayload() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        using Stream payload = document.OpenEmbeddedPayload(FirstPayloadIdentity, PayloadLength);
        return payload.Length;
    }

    internal static byte[] CreateFixture() {
        using var document = PowerPointDocument.Create();
        for (int index = 1; index <= SlideCount; index++) {
            document.AddSlide(PowerPointSlideLayoutType.Blank).AddTextBox("Embedded document benchmark " + index);
        }
        byte[] data = document.ToBytes();
        using var stream = new MemoryStream();
        stream.Write(data, 0, data.Length);
        stream.Position = 0;
        using (var package = Package.Open(stream, FileMode.Open, FileAccess.ReadWrite)) {
            var generator = new Random(3706);
            var payload = new byte[PayloadLength];
            for (int index = 1; index <= SlideCount; index++) {
                Uri slide = PartUri("/ppt/slides/slide" + index + ".xml");
                Uri target = PartUri("/ppt/embeddings/payload-" + index + ".bin");
                generator.NextBytes(payload);
                using (Stream bytes = package.CreatePart(target, "application/octet-stream", CompressionOption.Normal).GetStream()) {
                    bytes.Write(payload, 0, payload.Length);
                }
                package.GetPart(slide).CreateRelationship(target, TargetMode.Internal, PackageRelationship, PayloadRelationshipId);
            }
        }
        return stream.ToArray();
    }

    private static Uri PartUri(string path) => PackUriHelper.CreatePartUri(new Uri(path, UriKind.Relative));
}
