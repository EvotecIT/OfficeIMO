using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Security.Benchmarks;

[MemoryDiagnoser]
public class SecurityCmsSignBenchmarks {
    private SecurityCmsBenchmarkFixture _fixture = null!;

    [ParamsSource(nameof(CorpusScales))]
    public string Scale { get; set; } = string.Empty;

    public IEnumerable<string> CorpusScales() => SecurityCmsBenchmarkCorpus.Scales;

    [GlobalSetup]
    public void Setup() {
        _fixture = SecurityCmsBenchmarkCorpus.Create(Scale);
        SecurityCmsBenchmarkValidation.Validate(_fixture);
    }

    [GlobalCleanup]
    public void Cleanup() => _fixture.Dispose();

    [Benchmark(Baseline = true, Description = "OfficeIMO detached CMS sign")]
    public byte[] OfficeIMO() => SecurityCmsBenchmarkValidation.SignOffice(_fixture);

    [Benchmark(Description = ".NET detached CMS sign")]
    public byte[] Platform() => SecurityCmsBenchmarkValidation.SignPlatform(_fixture);
}

[MemoryDiagnoser]
public class SecurityCmsVerifyBenchmarks {
    private SecurityCmsBenchmarkFixture _fixture = null!;
    private byte[] _signature = null!;

    [ParamsSource(nameof(CorpusScales))]
    public string Scale { get; set; } = string.Empty;

    [Params("OfficeIMO", "Platform")]
    public string SignatureProducer { get; set; } = string.Empty;

    public IEnumerable<string> CorpusScales() => SecurityCmsBenchmarkCorpus.Scales;

    [GlobalSetup]
    public void Setup() {
        _fixture = SecurityCmsBenchmarkCorpus.Create(Scale);
        SecurityCmsValidationSnapshot snapshot = SecurityCmsBenchmarkValidation.Validate(_fixture);
        _signature = string.Equals(SignatureProducer, "OfficeIMO", StringComparison.Ordinal)
            ? snapshot.OfficeSignature
            : snapshot.PlatformSignature;
    }

    [GlobalCleanup]
    public void Cleanup() => _fixture.Dispose();

    [Benchmark(Baseline = true, Description = "OfficeIMO detached CMS verify")]
    public CmsVerificationResult OfficeIMO() =>
        SecurityCmsBenchmarkValidation.VerifyOffice(_signature, _fixture.Content);

    [Benchmark(Description = ".NET detached CMS verify")]
    public PlatformCmsVerificationSnapshot Platform() =>
        SecurityCmsBenchmarkValidation.VerifyPlatform(_signature, _fixture.Content);
}
