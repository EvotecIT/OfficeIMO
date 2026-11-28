using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointVbaSourceBenchmarks {
    private const string Module = "Helpers";
    internal const string Original = "' original VBA module";
    internal const string Edited = "' edited VBA module";
    private byte[] _fixture = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? path = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_VBA_SOURCE_FIXTURE");
        _fixture = path == null ? CreateFixture() : File.ReadAllBytes(path);
        if (!ReadSource().Contains(Original, StringComparison.Ordinal)) {
            throw new InvalidOperationException("The detached VBA source did not retain the original module.");
        }
        using var result = Load(EditSourceAndSave());
        if (result.InspectVbaProjectMetadata()!.HasSignature ||
            !result.ReadVbaProject()!.Modules.Single().Source.Contains(Edited, StringComparison.Ordinal)) {
            throw new InvalidOperationException("The saved deck did not retain the edited VBA source and remove its signature.");
        }
    }

    [Benchmark]
    public string ReadSource() {
        using var document = Load(_fixture);
        return document.ReadVbaProject()!.Modules.Single().Source;
    }

    [Benchmark]
    public byte[] EditSourceAndSave() {
        using var document = Load(_fixture);
        EditSource(document);
        return document.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Pptm });
    }

    internal static void EditSource(PowerPointDocument document) {
        OfficeVbaProject project = document.ReadVbaProject()!;
        project.SetModuleSource(Module, Edited);
        document.SetVbaProject(project, new OfficeVbaWriteOptions { AllowSignatureRemoval = true });
    }

    internal static byte[] CreateProject() {
        OfficeVbaProject project = OfficeVbaProject.Create("PowerPointSourceFixture");
        project.AddModule(Module, Original);
        return project.Write().GetBytes();
    }

    internal static byte[] CreateFixture() => PowerPointVbaReadBenchmarks.AttachProject(
        PowerPointModernSubtypeBenchmarks.CreateFixture(PresentationDocumentType.Presentation),
        CreateProject(), PresentationDocumentType.MacroEnabledPresentation);

    private static PowerPointDocument Load(byte[] bytes) => PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
}
