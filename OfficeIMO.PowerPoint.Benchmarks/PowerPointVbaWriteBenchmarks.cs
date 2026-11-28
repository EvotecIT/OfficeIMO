using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointVbaWriteBenchmarks {
    private byte[] _fixture = Array.Empty<byte>();
    private byte[] _project = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? location = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_VBA_WRITE_FIXTURE");
        if (location == null) {
            _project = PowerPointVbaReadBenchmarks.ReadProject(PowerPointVbaReadBenchmarks.FindProjectFixture());
            _fixture = PowerPointVbaReadBenchmarks.CreateFixture(_project);
        } else {
            _fixture = File.ReadAllBytes(location);
            using var source = Load();
            _project = source.GetVbaProjectBytes()!;
            if (!source.InspectVbaProjectMetadata()!.HasSignature) throw new InvalidOperationException("The fixture has no VBA signature part.");
        }

        using var replaced = PowerPointDocument.Load(new MemoryStream(ReplaceProjectAndSave(), writable: false));
        using var removed = PowerPointDocument.Load(new MemoryStream(RemoveProjectAndSave(), writable: false));
        if (replaced.SourceFormatDescriptor.Extension != ".pptm" || replaced.InspectVbaProjectMetadata()!.HasSignature ||
            !replaced.GetVbaProjectBytes()!.SequenceEqual(_project) || removed.HasVbaProject ||
            removed.SourceFormatDescriptor.Extension != ".pptm") {
            throw new InvalidOperationException("The saved presentations did not preserve the expected VBA mutation or subtype.");
        }
    }

    [Benchmark]
    public byte[] ReplaceProjectAndSave() {
        using PowerPointDocument document = Load();
        document.SetVbaProject(_project);
        return Save(document);
    }

    [Benchmark]
    public byte[] RemoveProjectAndSave() {
        using PowerPointDocument document = Load();
        document.RemoveVbaProject();
        return Save(document);
    }

    private PowerPointDocument Load() => PowerPointDocument.Load(new MemoryStream(_fixture, writable: false));

    private static byte[] Save(PowerPointDocument document) =>
        document.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Pptm });
}
