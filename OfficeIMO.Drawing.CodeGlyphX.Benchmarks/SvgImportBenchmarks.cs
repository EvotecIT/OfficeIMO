using System;
using System.Collections.Generic;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Drawing;

[MemoryDiagnoser(displayGenColumns: false)]
[RankColumn]
public class SvgImportBenchmarks {
    private byte[] _svg = Array.Empty<byte>();

    [ParamsSource(nameof(ScenarioNames))]
    public string Scenario { get; set; } = string.Empty;

    public IEnumerable<string> ScenarioNames => SvgScenarioFactory.Names;

    [GlobalSetup]
    public void Setup() {
        _svg = SvgScenarioFactory.Create(Scenario).Svg;
    }

    [Benchmark]
    public OfficeDrawing ImportSvg() {
        if (!OfficeSvgDrawingReader.TryRead(_svg, out OfficeDrawing? drawing, out _) || drawing is null) {
            throw new InvalidOperationException($"Could not import benchmark scenario '{Scenario}'.");
        }
        return drawing;
    }
}
