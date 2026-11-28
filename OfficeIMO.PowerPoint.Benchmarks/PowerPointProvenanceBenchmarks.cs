using System.IO.Compression;
using System.IO.Packaging;
using System.Text;
using System.Xml.Linq;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Provenance;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointProvenanceBenchmarks {
    private const string SignatureRelationship = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/";
    private const string SignatureType = "application/vnd.openxmlformats-package.digital-signature-";
    private const string Origin = "/_xmlsignatures/origin.sigs";
    private const string Signature = "/_xmlsignatures/sig1.xml";
    private const string Manifest = "/META-INF/content_credential.c2pa";
    private static readonly OfficeProvenanceRemovalOptions Removal = new() {
        RequireStructurallyValidCarrier = false,
        SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
    };
    private byte[] _input = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? path = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_PROVENANCE_FIXTURE");
        _input = string.IsNullOrEmpty(path) ? CreateFixture() : File.ReadAllBytes(path);
        OfficeProvenanceRemovalResult result = PowerPointDocument.RemoveProvenance(_input, options: Removal);
        using var output = new MemoryStream(result.ToArray(), writable: false);
        using var archive = new ZipArchive(output, ZipArchiveMode.Read);
        using Stream root = archive.GetEntry("_rels/.rels")!.Open();
        bool hasSignatureRelationship = XElement.Load(root).Elements().Any(element =>
            (string?)element.Attribute("Type") == SignatureRelationship + "origin");
        if (!result.WereInvalidatedSignaturesRemoved || InspectProvenance().Evidence.Count != 1 ||
            archive.GetEntry(Signature.TrimStart('/')) != null || archive.GetEntry(Manifest.TrimStart('/')) != null ||
            hasSignatureRelationship) {
            throw new InvalidOperationException("The benchmark did not inspect the carrier and remove the manifest and signature.");
        }
    }

    [Benchmark]
    public OfficeProvenanceReport InspectProvenance() => PowerPointDocument.InspectProvenance(_input);

    [Benchmark]
    public byte[] RemoveManifestAndSignatures() => PowerPointDocument.RemoveProvenance(_input, options: Removal).ToArray();

    internal static byte[] CreateFixture() {
        using var document = PowerPointDocument.Create();
        for (int index = 0; index < 20; index++) {
            document.AddSlide(PowerPointSlideLayoutType.Blank).AddTextBox("Presentation provenance benchmark");
        }
        return WithManifestAndSignatures(document.ToBytes());
    }

    internal static byte[] WithManifestAndSignatures(byte[] presentation) {
        using var destination = new MemoryStream();
        destination.Write(presentation, 0, presentation.Length);
        destination.Position = 0;
        using (var package = Package.Open(destination, FileMode.Open, FileAccess.ReadWrite)) {
            PackagePart manifest = package.CreatePart(PartUri(Manifest), "application/c2pa", CompressionOption.Normal);
            using (Stream data = manifest.GetStream(FileMode.Create, FileAccess.Write)) {
                data.WriteByte(1);
            }
            PackagePart origin = package.CreatePart(PartUri(Origin), SignatureType + "origin", CompressionOption.Normal);
            package.CreateRelationship(PartUri(Origin), TargetMode.Internal, SignatureRelationship + "origin", "provenanceOrigin");
            PackagePart signature = package.CreatePart(PartUri(Signature), SignatureType + "xmlsignature+xml", CompressionOption.Normal);
            origin.CreateRelationship(PartUri(Signature), TargetMode.Internal, SignatureRelationship + "signature", "provenanceSignature");
            var sample = new byte[16 * 1024];
            new Random(3706).NextBytes(sample);
            byte[] xml = Encoding.UTF8.GetBytes("<Signature xmlns='http://www.w3.org/2000/09/xmldsig#'><Object>" +
                Convert.ToBase64String(sample) + "</Object></Signature>");
            using Stream output = signature.GetStream(FileMode.Create, FileAccess.Write);
            output.Write(xml, 0, xml.Length);
        }
        return destination.ToArray();
    }

    private static Uri PartUri(string path) => PackUriHelper.CreatePartUri(new Uri(path, UriKind.Relative));
}
