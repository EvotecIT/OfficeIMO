using System;
using System.IO;
using System.Linq;
using ChartForgeX.VisualArtifacts;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

public sealed partial class OfficeVisioVisualIntegrationTests {
    [Fact]
    public void PreservedMultilineTitleFitsClearCorridorWithoutMovingPreparedGeometry() {
        var envelope = MultilineTitlePlacementEnvelope();
        var result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { PixelsPerInch = 96 });
        Assert.DoesNotContain(result.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.TitleNotProjected);
        void Check(VisioPage page) {
            var title = page.Shapes.Single(shape => shape.Id == "cfx-title");
            Assert.Equal("Service delivery\nThree editable services · prepared routes", title.Text.Replace("\r\n", "\n"));
            Assert.Equal(22D, title.TextStyle!.Size!.Value, 6);
            Assert.True(title.Height > 2D * 22D / 72D);
            var bounds = title.GetShapeBounds();
            Assert.True(bounds.Top <= page.Height);
            Assert.True(bounds.Bottom >= 0);
            Assert.All(page.Shapes.Where(shape => shape.Id == "api" || shape.Id == "database"),
                shape => Assert.True(bounds.Bottom >= shape.GetShapeBounds().Top + 0.08D - 0.000001D));
            var api = page.Shapes.Single(shape => shape.Id == "api");
            Assert.Equal(106D / 96D, api.PinX, 6);
            Assert.Equal(110D / 96D, api.PinY, 6);
            Assert.Equal(148D / 96D, api.Width, 6);
            Assert.Equal(72D / 96D, api.Height, 6);
            var connector = Assert.Single(page.Connectors);
            Assert.True(bounds.Bottom >= connector.GetConnectorContentBounds().Top + 0.08D - 0.000001D);
            Assert.Collection(connector.Waypoints,
                point => { Assert.Equal(300D / 96D, point.X, 6); Assert.Equal(128D / 96D, point.Y, 6); },
                point => { Assert.Equal(300D / 96D, point.X, 6); Assert.Equal(110D / 96D, point.Y, 6); });
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
    public void PreservedMultilineTitleReportsOmissionWhenMeasuredHeightCrowdsNodes() {
        var envelope = MultilineTitlePlacementEnvelope();
        envelope.Nodes[0].Y = 70;
        envelope.Edges[0].ResolvedRoute[0].Y = 88;
        var result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { PixelsPerInch = 96 });
        Assert.DoesNotContain(result.Page.Shapes, shape => shape.Id == "cfx-title");
        Assert.Contains(result.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.TitleNotProjected);
        Assert.Equal(254D / 96D, result.Page.Shapes.Single(shape => shape.Id == "api").PinY, 6);
        var rejected = Assert.Throws<OfficeVisioVisualFidelityException>(() => envelope.ToOfficeVisio(
            new OfficeVisioVisualOptions { PixelsPerInch = 96, RequireLossless = true }));
        Assert.Contains(rejected.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.TitleNotProjected);
    }

    private static VisualArtifactInterchangeEnvelope MultilineTitlePlacementEnvelope() {
        var envelope = PlacementEnvelope();
        envelope.Title = "Service delivery";
        envelope.Subtitle = "Three editable services · prepared routes";
        envelope.Width = 600;
        envelope.Height = 360;
        var api = envelope.Nodes[0];
        api.X = 32; api.Y = 214; api.Width = 148; api.Height = 72;
        var database = envelope.Nodes[1];
        database.X = 420; database.Y = 214; database.Width = 148; database.Height = 72;
        var edge = envelope.Edges[0];
        edge.Label = "prepared route";
        edge.ResolvedRoute.Add(new VisualArtifactInterchangePoint { X = 180, Y = 232 });
        edge.ResolvedRoute.Add(new VisualArtifactInterchangePoint { X = 300, Y = 232 });
        edge.ResolvedRoute.Add(new VisualArtifactInterchangePoint { X = 300, Y = 250 });
        edge.ResolvedRoute.Add(new VisualArtifactInterchangePoint { X = 420, Y = 250 });
        return envelope;
    }

}
