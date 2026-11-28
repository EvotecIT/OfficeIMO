using System.IO.Compression;
using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointVbaReadBenchmarks {
    private byte[] _presentation = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_VBA_FIXTURE");
        _presentation = fixture == null ? CreateFixture(ReadProject(FindProjectFixture())) : File.ReadAllBytes(fixture);
        OfficeVbaProjectInfo descriptor = ReadProjectMetadata();
        OfficeVbaProjectInfo inspected = ReadProjectDetails();
        if (descriptor.HasKnownLength || descriptor.ModuleInspectionRequested || !descriptor.HasSignature ||
            !inspected.ModuleInspectionRequested || inspected.ModuleNames.Count == 0 || inspected.Sha256?.Length != 64 ||
            inspected.Length != ReadProjectBytes().LongLength) {
            throw new InvalidOperationException("The benchmark fixture did not preserve its project and signature metadata.");
        }
    }

    [Benchmark]
    public OfficeVbaProjectInfo ReadProjectMetadata() {
        using PowerPointDocument presentation = Load();
        return presentation.InspectVbaProjectMetadata()!;
    }

    [Benchmark]
    public OfficeVbaProjectInfo ReadProjectDetails() {
        using PowerPointDocument presentation = Load();
        return presentation.InspectVbaProject(includeSha256: true)!;
    }

    [Benchmark]
    public byte[] ReadProjectBytes() {
        using PowerPointDocument presentation = Load();
        return presentation.GetVbaProjectBytes()!;
    }

    private PowerPointDocument Load() => PowerPointDocument.Load(new MemoryStream(_presentation, writable: false),
        new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });

    internal static byte[] CreateFixture(byte[] project) {
        using var presentation = PowerPointDocument.Create();
        for (int i = 0; i < 24; i++) {
            presentation.AddSlide(PowerPointSlideLayoutType.Blank).AddTextBox("VBA metadata benchmark");
        }
        return AttachProject(presentation.ToBytes(), project, PresentationDocumentType.MacroEnabledPresentation);
    }

    internal static byte[] AttachProject(byte[] bytes, byte[] project, PresentationDocumentType documentType) {
        using var destination = new MemoryStream();
        destination.Write(bytes, 0, bytes.Length);
        destination.Position = 0;
        using (var package = PresentationDocument.Open(destination, true)) {
            package.ChangeDocumentType(documentType);
            VbaProjectPart part = package.PresentationPart!.AddNewPart<VbaProjectPart>();
            using var source = new MemoryStream(project, writable: false);
            part.FeedData(source);
            var signature = part.AddExtendedPart("http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature",
                "application/vnd.ms-office.vbaProjectSignature", ".bin");
            using var payload = new MemoryStream(new byte[] { 7 }, writable: false);
            signature.FeedData(payload);
        }
        return destination.ToArray();
    }

    internal static string FindProjectFixture() {
        for (DirectoryInfo? location = new(AppContext.BaseDirectory); location != null; location = location.Parent) {
            string file = Path.Combine(location.FullName, "OfficeIMO.TestAssets", "Documents", "ExcelPreservationCorpus", "inert-macro-control.xlsm");
            if (File.Exists(file)) return file;
        }
        throw new FileNotFoundException("The inert VBA test asset is missing.");
    }

    internal static byte[] ReadProject(string path) {
        using var archive = new ZipArchive(File.OpenRead(path), ZipArchiveMode.Read);
        using Stream entry = archive.GetEntry("xl/vbaProject.bin")!.Open();
        using var result = new MemoryStream();
        entry.CopyTo(result);
        return result.ToArray();
    }
}
