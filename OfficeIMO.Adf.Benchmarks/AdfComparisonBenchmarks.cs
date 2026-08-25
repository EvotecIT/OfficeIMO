using System.Text.Json.Nodes;
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Adf.Benchmarks;

/// <summary>Measures the typed ADF parse cost against platform JSON-tree and typed-model floors.</summary>
[MemoryDiagnoser]
public class AdfParseBenchmarks {
    private string _json = string.Empty;

    [Params("Small", "Normal")]
    public string Scale { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup() {
        _json = AdfBenchmarkCorpus.Create(AdfBenchmarkCorpus.Get(Scale));
        AdfComparisonValidation.ValidateOfficeParse(_json, AdfComparisonWorkflows.ParseOfficeIMO(_json));
        AdfComparisonValidation.ValidatePlatformParse(_json, AdfComparisonWorkflows.ParsePlatform(_json));
        AdfComparisonValidation.ValidatePlatformTypedParse(_json, AdfComparisonWorkflows.ParsePlatformTyped(_json));
    }

    [Benchmark(Description = "System.Text.Json tree floor")]
    public JsonNode Platform() => AdfComparisonWorkflows.ParsePlatform(_json);

    [Benchmark(Baseline = true, Description = "System.Text.Json typed floor")]
    public PlatformAdfDocument PlatformTyped() => AdfComparisonWorkflows.ParsePlatformTyped(_json);

    [Benchmark(Description = "OfficeIMO typed ADF")]
    public AdfDocument OfficeIMO() => AdfComparisonWorkflows.ParseOfficeIMO(_json);
}

/// <summary>Measures a semantic JSON round trip against platform JSON-tree and typed-model floors.</summary>
[MemoryDiagnoser]
public class AdfRoundTripBenchmarks {
    private string _json = string.Empty;

    [Params("Small", "Normal")]
    public string Scale { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup() {
        _json = AdfBenchmarkCorpus.Create(AdfBenchmarkCorpus.Get(Scale));
        AdfComparisonValidation.Inspect(_json, AdfComparisonWorkflows.RoundTripOfficeIMO(_json), "OfficeIMO");
        AdfComparisonValidation.Inspect(_json, AdfComparisonWorkflows.RoundTripPlatform(_json), "System.Text.Json");
        AdfComparisonValidation.Inspect(_json, AdfComparisonWorkflows.RoundTripPlatformTyped(_json), "System.Text.Json typed model");
    }

    [Benchmark(Description = "System.Text.Json tree floor")]
    public string Platform() => AdfComparisonWorkflows.RoundTripPlatform(_json);

    [Benchmark(Baseline = true, Description = "System.Text.Json typed floor")]
    public string PlatformTyped() => AdfComparisonWorkflows.RoundTripPlatformTyped(_json);

    [Benchmark(Description = "OfficeIMO typed ADF")]
    public string OfficeIMO() => AdfComparisonWorkflows.RoundTripOfficeIMO(_json);
}
