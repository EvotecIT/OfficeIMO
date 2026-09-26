using System;
using System.IO;
using System.Linq;
using ChartForgeX.Topology;
using ChartForgeX.VisualArtifacts;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

public sealed partial class OfficeVisioVisualIntegrationTests {
    [Theory]
    [InlineData(0, OfficeVisioVisualDiagnosticCode.GeometryOutsidePage)]
    [InlineData(70, OfficeVisioVisualDiagnosticCode.TitleNotProjected)]
    public void PreservedLabelBoundsParticipateInFidelityAndTitleClearance(double y, OfficeVisioVisualDiagnosticCode code) {
        var envelope = PlacementEnvelope();
        envelope.Title = "Services";
        foreach (var node in envelope.Nodes) node.Y = y;
        envelope.Edges[0].Label = "Relationship";
        envelope.Edges[0].Topology!.Waypoints.Clear();
        var options = new OfficeVisioVisualOptions { PixelsPerInch = 100 };
        var result = envelope.ToOfficeVisio(options);
        Assert.Contains(result.Report.Diagnostics, item => item.Code == code);
        Assert.DoesNotContain(result.Page.Shapes, shape => shape.Text == "Services");
        options.RejectedDiagnostics.Add(code);
        Assert.Throws<OfficeVisioVisualFidelityException>(() => envelope.ToOfficeVisio(options));
    }

    [Fact]
    public void PreservedGroupCaptionStaysInsideTopBoundaryThroughSave() {
        var envelope = PlacementEnvelope();
        envelope.Groups.Add(new VisualArtifactInterchangeGroup {
            Id = "zone", Label = "Top boundary", Kind = "TopologyGroup",
            Role = VisualArtifactInterchangeGroupRole.TopologyGroup,
            X = 0, Y = 0, Width = 400, Height = 250,
            Topology = new VisualArtifactInterchangeTopologyGroup()
        });
        envelope.Nodes[0].GroupId = "zone";
        var result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { PixelsPerInch = 100 });
        void Check(VisioPage page) {
            var caption = page.Shapes.Single(shape => shape.Text == "Top boundary");
            Assert.InRange(caption.PinY + caption.Height / 2, 0, page.Height);
            Assert.InRange(caption.PinX - caption.Width / 2, 0, 4);
            Assert.InRange(caption.PinX + caption.Width / 2, 0, 4);
        }
        Check(result.Page);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            result.Document.Save(path);
            Check(VisioDocument.Load(path).Pages[0]);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void GeneratedSelfRouteReservesItsHeaderSpace() {
        var envelope = PlacementEnvelope();
        envelope.Title = "Services";
        envelope.Nodes[0].Y = 80;
        var edge = envelope.Edges[0];
        edge.TargetId = "api";
        edge.TargetPortId = edge.SourcePortId;
        edge.Topology!.Waypoints.Clear();
        var result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { PixelsPerInch = 96 });
        Assert.Contains(result.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.TitleNotProjected);
        Assert.DoesNotContain(result.Page.Shapes, shape => shape.Text == "Services");
    }

    [Fact]
    public void ThemeLossesIdentifyTheArtifactInsteadOfANode() {
        var chart = TopologyChart.Create().WithTheme(TopologyTheme.Light()).AddNode("service", "Service", 100, 100);
        chart.Theme!.Foreground = "rgba(1,2,3,0.5)";
        var envelope = chart.ToVisualArtifact().ToInterchangeEnvelope();
        var result = envelope.ToOfficeVisio();
        var losses = result.Report.Diagnostics.Where(item => item.Code == OfficeVisioVisualDiagnosticCode.ColorNotProjected).ToArray();
        Assert.NotEmpty(losses);
        Assert.All(losses, item => {
            Assert.Equal(OfficeVisioVisualEntityKind.Artifact, item.EntityKind);
            Assert.Equal(envelope.Id, item.EntityId);
        });
    }

    [Fact]
    public void RejectedTypedProjectionIncludesAllRequestedFeatureLosses() {
        var chart = TopologyChart.Create().WithTheme(TopologyTheme.Light()).AddAutoNode("service", "Service");
        var render = new VisualArtifactRenderOptions();
        render.Watermarks.Add(VisualWatermark.FromText("CONFIDENTIAL"));
        var error = Assert.Throws<OfficeVisioVisualFidelityException>(() => chart.ToVisualArtifact().ToOfficeVisio(
            new OfficeVisioVisualOptions { RequireLossless = true }, render));
        Assert.Contains(error.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.WatermarkNotProjected);
        Assert.Contains(error.Report.Diagnostics, item => item.Code != OfficeVisioVisualDiagnosticCode.WatermarkNotProjected);
    }
}
