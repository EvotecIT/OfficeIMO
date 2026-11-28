using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointModernSubtypeSaveBenchmarks {
    private const string Edit = "Saved in the original subtype";
    private byte[] _template = Array.Empty<byte>();
    private byte[] _macro = Array.Empty<byte>();
    private string _destination = string.Empty;

    [GlobalSetup]
    public void Setup() {
        string? directory = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_SUBTYPE_SAVE_FIXTURES");
        _template = directory == null ? PowerPointModernSubtypeBenchmarks.CreateFixture(PresentationDocumentType.Template)
            : File.ReadAllBytes(Path.Combine(directory, "fixture.potx"));
        _macro = directory == null ? PowerPointVbaReadBenchmarks.CreateFixture(PowerPointVbaReadBenchmarks.ReadProject(
            PowerPointVbaReadBenchmarks.FindProjectFixture())) : File.ReadAllBytes(Path.Combine(directory, "fixture.pptm"));
        _destination = Path.Combine(Path.GetTempPath(), "OfficeIMO-PowerPointSubtype-" + Guid.NewGuid().ToString("N") + ".potx");

        SaveTemplateCopy();
        using var template = PowerPointDocument.Load(_destination);
        if (template.SourceFormatDescriptor.Extension != ".potx" || !template.Slides[0].TextBoxes.Any(box => box.Text == Edit)) {
            throw new InvalidOperationException("The edited template failed to preserve its subtype.");
        }
        using var originalMacro = Load(_macro);
        using var savedMacro = Load(SaveMacroPresentation());
        if (savedMacro.SourceFormatDescriptor.Extension != ".pptm" || !savedMacro.InspectVbaProjectMetadata()!.HasSignature ||
            !savedMacro.GetVbaProjectBytes()!.SequenceEqual(originalMacro.GetVbaProjectBytes()!) ||
            !savedMacro.Slides[0].TextBoxes.Any(box => box.Text == Edit)) {
            throw new InvalidOperationException("The edited macro presentation failed to preserve its subtype or project.");
        }
    }

    [GlobalCleanup]
    public void Cleanup() {
        if (File.Exists(_destination)) File.Delete(_destination);
    }

    [Benchmark]
    public long SaveTemplateCopy() {
        using PowerPointDocument document = Load(_template);
        document.Slides[0].AddTextBox(Edit);
        document.SaveCopy(_destination);
        return new FileInfo(_destination).Length;
    }

    [Benchmark]
    public byte[] SaveMacroPresentation() {
        using PowerPointDocument document = Load(_macro);
        document.Slides[0].AddTextBox(Edit);
        return document.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Pptm });
    }

    private static PowerPointDocument Load(byte[] bytes) => PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
}
