using System;
using System.Linq;
using ChartForgeX.Topology;
using ChartForgeX.VisualArtifacts;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

public sealed partial class OfficeVisioVisualIntegrationTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void VisioProjectionReportsWatermarksAcrossTypedEnvelopeAndJsonIngress(int ingress) {
        var artifact = TopologyChart.Create().AddAutoNode("service", "Service")
            .ToVisualArtifact().WithWatermarks(VisualWatermark.FromText("CONFIDENTIAL"));
        Func<OfficeVisioVisualOptions, OfficeVisioVisualConversionResult> project = ingress switch {
            0 => options => artifact.ToOfficeVisio(options),
            1 => options => artifact.ToInterchangeEnvelope().ToOfficeVisio(options),
            _ => options => artifact.ToInterchangeUtf8Json().ToOfficeVisio(options)
        };

        var result = project(new OfficeVisioVisualOptions());
        var warning = Assert.Single(result.Report.Diagnostics, item =>
            item.Code == OfficeVisioVisualDiagnosticCode.WatermarkNotProjected);
        Assert.Equal(artifact.Id, warning.EntityId);
        Assert.Contains("SVG or PNG", warning.Message);
        var strict = new OfficeVisioVisualOptions();
        strict.RejectedDiagnostics.Add(OfficeVisioVisualDiagnosticCode.WatermarkNotProjected);
        var error = Assert.Throws<OfficeVisioVisualFidelityException>(() => project(strict));
        Assert.Contains(error.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.WatermarkNotProjected);
    }
}
