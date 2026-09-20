using System;
using System.IO;
using System.Linq;
using ChartForgeX.Topology;
using ChartForgeX.Primitives;
using ChartForgeX.VisualArtifacts;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

public sealed partial class OfficeVisioVisualIntegrationTests {
    [Fact]
    public void PreparedBoundsAndNamedAttachmentsSurviveSaveAndLoad() {
        var envelope = PlacementEnvelope();
        var result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { PixelsPerInch = 100 });
        Assert.DoesNotContain(result.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.LayoutRecomputed);
        void Check(VisioPage page) {
            Assert.Equal(10, page.Width, 6);
            Assert.Equal(8, page.Height, 6);
            var node = page.Shapes.Single(item => item.Id == "api");
            Assert.Equal(2, node.PinX, 6);
            Assert.Equal(6.5, node.PinY, 6);
            Assert.Equal(2, node.Width, 6);
            Assert.Equal(1, node.Height, 6);
            var connector = Assert.Single(page.Connectors);
            Assert.Equal(VisioConnectorRerouteBehavior.Never, connector.RerouteBehavior);
            Assert.Equal(2, connector.FromConnectionPoint!.X, 6);
            Assert.Equal(.75, connector.FromConnectionPoint.Y, 6);
            Assert.Contains(connector.Waypoints, point => Math.Abs(point.X - 4) < .000001 && Math.Abs(point.Y - 6.75) < .000001);
        }
        Check(result.Page);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            result.Document.Save(path);
            Assert.Empty(VisioValidator.Validate(path));
            Check(VisioDocument.Load(path).Pages[0]);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void ExplicitReflowAndFidelityPolicyAreHonored() {
        var envelope = PlacementEnvelope();
        var reflow = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { LayoutMode = OfficeVisioVisualLayoutMode.Reflow });
        Assert.Contains(reflow.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.LayoutRecomputed);
        var options = new OfficeVisioVisualOptions { LayoutMode = OfficeVisioVisualLayoutMode.Reflow };
        options.RejectedDiagnostics.Add(OfficeVisioVisualDiagnosticCode.LayoutRecomputed);
        var error = Assert.Throws<OfficeVisioVisualFidelityException>(() => envelope.ToOfficeVisio(options));
        Assert.Contains(error.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.LayoutRecomputed);
        envelope.Nodes[0].Width = null;
        Assert.Throws<NotSupportedException>(() => envelope.ToOfficeVisio(new OfficeVisioVisualOptions { LayoutMode = OfficeVisioVisualLayoutMode.Preserve }));
    }

    [Fact]
    public void PreserveRejectsSequenceInsteadOfSilentlyReflowing() {
        Assert.Throws<NotSupportedException>(() => SequenceEnvelope("sequence").ToOfficeVisio(
            new OfficeVisioVisualOptions { LayoutMode = OfficeVisioVisualLayoutMode.Preserve }));
    }

    [Fact]
    public void BookPreservesInputOrderAndDistinctPagesThroughSave() {
        var first = PlacementEnvelope(); var second = PlacementEnvelope();
        first.Title = second.Title = "Services";
        var book = new[] { first, second }.ToOfficeVisioBook();
        Assert.Equal(new[] { "Services", "Services (2)" }, book.Document.Pages.Select(page => page.Name));
        Assert.All(book.Pages, page => Assert.Same(book.Document, page.Document));
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            book.Document.Save(path);
            Assert.Empty(VisioValidator.Validate(path));
            Assert.Equal(2, VisioDocument.Load(path).Pages.Count);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void PreservedOverflowIsReportedAndCanBeRejected() {
        var envelope = PlacementEnvelope();
        envelope.Nodes[0].X = -10;
        var result = envelope.ToOfficeVisio();
        Assert.Contains(result.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.GeometryOutsidePage && item.EntityId == "api");
        var options = new OfficeVisioVisualOptions();
        options.RejectedDiagnostics.Add(OfficeVisioVisualDiagnosticCode.GeometryOutsidePage);
        Assert.Throws<OfficeVisioVisualFidelityException>(() => envelope.ToOfficeVisio(options));
    }

    [Fact]
    public void SelfRouteKeepsExplicitPortsAndLosslessPolicyRejectsNormalization() {
        var envelope = PlacementEnvelope();
        var edge = envelope.Edges[0];
        edge.TargetId = "api";
        edge.TargetPortId = "out";
        edge.Topology!.Waypoints.Clear();
        var result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { PixelsPerInch = 100 });
        var connector = Assert.Single(result.Page.Connectors);
        Assert.Equal(.75, connector.FromConnectionPoint!.Y, 6);
        Assert.Equal(.75, connector.ToConnectionPoint!.Y, 6);
        Assert.True(connector.Waypoints.Count >= 3);
        Assert.Throws<OfficeVisioVisualFidelityException>(() => envelope.ToOfficeVisio(new OfficeVisioVisualOptions { RequireLossless = true }));
    }

    [Fact]
    public void NativeGraphUsesSourceCardColorsAndAllowsExplicitThemeOverride() {
        var chart = TopologyChart.Create().WithTheme(TopologyTheme.Light()).AddNode("a", "Service", 100, 100);
        chart.Theme!.Card = "#F0F4FA";
        chart.Theme.Foreground = "#14243A";
        var envelope = chart.ToVisualArtifact().ToInterchangeEnvelope();
        var result = envelope.ToOfficeVisio();
        var shape = result.Page.Shapes.Single(item => item.Id == "a");
        Assert.Equal(OfficeIMO.Drawing.OfficeColor.FromRgb(240, 244, 250), shape.FillColor);
        Assert.Equal(OfficeIMO.Drawing.OfficeColor.FromRgb(20, 36, 58), shape.TextStyle!.Color);
        Assert.Equal("Arial", shape.TextStyle.FontFamily);
        var custom = VisioStyleTheme.Technical();
        var overridden = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { NativeTheme = custom });
        Assert.Equal(custom.Primary.FillColor, overridden.Page.Shapes.Single(item => item.Id == "a").FillColor);
    }

    private static VisualArtifactInterchangeEnvelope PlacementEnvelope() {
        var envelope = TopologyEnvelope("placed");
        envelope.Width = 1000; envelope.Height = 800;
        var api = TopologyNode("api", "API");
        api.X = 100; api.Y = 100; api.Width = 200; api.Height = 100;
        api.Ports.Add(new VisualArtifactInterchangePort { Id = "out", Side = TopologyEdgePort.Right, Offset = .25 });
        var database = TopologyNode("database", "Database");
        database.X = 600; database.Y = 400; database.Width = 200; database.Height = 100;
        envelope.Nodes.Add(api); envelope.Nodes.Add(database);
        var edge = TopologyEdge("link", VisualLinkDirection.Forward);
        edge.SourcePortId = "out";
        edge.Topology!.Routing = TopologyEdgeRouting.Straight;
        edge.Topology.Waypoints.Add(new VisualArtifactInterchangePoint { X = 400, Y = 125 });
        edge.Topology.Waypoints.Add(new VisualArtifactInterchangePoint { X = 400, Y = 450 });
        envelope.Edges.Add(edge);
        return envelope;
    }
}
