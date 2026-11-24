using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Scaling.Benchmarks;

/// <summary>
/// Measures each format lane at 10, 100 and 1,000 units. The per-unit columns
/// divide the mean time and allocation by <see cref="Units"/>, so a linear
/// operation shows a flat per-unit cost across sizes and a quadratic one grows
/// with the size.
/// </summary>
[MemoryDiagnoser]
[Config(typeof(PerUnitConfig))]
public class FormatScalingBenchmarks {
    private ScalingLane _lane = null!;
    private object? _result;

    public static IEnumerable<string> LaneNames => ScalingLanes.Create().Select(lane => lane.Name);

    [ParamsSource(nameof(LaneNames))]
    public string Lane { get; set; } = string.Empty;

    [Params(10, 100, 1000)]
    public int Units { get; set; }

    [GlobalSetup]
    public void Setup() {
        _lane = ScalingLanes.Create().Single(lane => lane.Name == Lane);
        _lane.Setup(Units);
    }

    [Benchmark]
    public object? Run() => _result = _lane.Run();

    [GlobalCleanup]
    public void Cleanup() => _lane.Validate(_result);
}
