using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointModernSubtypeConversionBenchmarks {
    private const string Edit = "Saved from another subtype";
    private byte[] _plain = Array.Empty<byte>();
    private byte[] _template = Array.Empty<byte>();
    private byte[] _macro = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? directory = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_SUBTYPE_CONVERSION_FIXTURES");
        if (directory == null) {
            _plain = PowerPointModernSubtypeBenchmarks.CreateFixture(PresentationDocumentType.Presentation);
            _template = PowerPointModernSubtypeBenchmarks.ChangeSubtype(_plain, PresentationDocumentType.Template);
            byte[] project = PowerPointVbaReadBenchmarks.ReadProject(PowerPointVbaReadBenchmarks.FindProjectFixture());
            _macro = PowerPointVbaReadBenchmarks.AttachProject(_plain, project, PresentationDocumentType.MacroEnabledPresentation);
        } else {
            _plain = File.ReadAllBytes(Path.Combine(directory, "fixture.pptx"));
            _template = File.ReadAllBytes(Path.Combine(directory, "fixture.potx"));
            _macro = File.ReadAllBytes(Path.Combine(directory, "fixture.pptm"));
        }

        CheckOutput(SaveTemplateAsMacroPresentation(), ".pptm");
        CheckOutput(SaveSignedMacroAsPptx(), ".pptx");
        CheckOutput(SavePptxAsPptx(), ".pptx");
    }

    [Benchmark] public byte[] SaveTemplateAsMacroPresentation() => Save(_template, PowerPointFileFormat.Pptm);

    [Benchmark] public byte[] SaveSignedMacroAsPptx() => Save(_macro, PowerPointFileFormat.Pptx);

    [Benchmark] public byte[] SavePptxAsPptx() => Save(_plain, PowerPointFileFormat.Pptx);

    private static byte[] Save(byte[] source, PowerPointFileFormat format) {
        using var document = PowerPointDocument.Load(new MemoryStream(source, writable: false));
        document.Slides[0].AddTextBox(Edit);
        return document.ToBytes(new PowerPointSaveOptions { Format = format });
    }

    private static void CheckOutput(byte[] bytes, string extension) {
        using var saved = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        if (saved.SourceFormatDescriptor.Extension != extension || saved.HasVbaProject ||
            !saved.Slides[0].TextBoxes.Any(box => box.Text == Edit)) {
            throw new InvalidOperationException("The saved deck did not have the expected subtype and slide edit.");
        }
    }
}
