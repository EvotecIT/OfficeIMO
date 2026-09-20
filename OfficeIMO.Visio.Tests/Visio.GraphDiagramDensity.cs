using System;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioGraphDiagramDensityTests {
    [Fact]
    public void PreservedPageFitReservesLegendAboveContent() {
        var document = VisioDocument.Create().GraphDiagram("Legend", graph =>
            graph.PageSize(4, 4).PreserveLayout().Legend().Import(new[] {
                new VisioGraphNodeRecord("a", "A") { Placement = new VisioGraphPlacement(1, 12, 1, 1) }
            }, Array.Empty<VisioGraphEdgeRecord>()));
        var page = document.Pages[0];
        var node = page.Shapes.Single(shape => shape.Id == "a");
        Assert.All(page.Shapes.Where(shape => shape.Id != "a"), shape =>
            Assert.True(shape.PinY - shape.Height / 2 >= node.PinY + node.Height / 2));
    }

    [Theory]
    [InlineData(VisioMeasurementUnit.Inches)]
    [InlineData(VisioMeasurementUnit.Centimeters)]
    public void PreservedPageFitIncludesRouteBendsAndLabels(VisioMeasurementUnit unit) {
        var nodes = new[] {
            new VisioGraphNodeRecord("a", "A") { Placement = new VisioGraphPlacement(1, 1, 1, 1) },
            new VisioGraphNodeRecord("b", "B") { Placement = new VisioGraphPlacement(3, 1, 1, 1) }
        };
        var edges = new[] { new VisioGraphEdgeRecord("ab", "a", "b") {
            Label = "An extended route label",
            Route = new VisioGraphRoute(new[] {
                new VisioConnectorWaypoint(1.5, 1), new VisioConnectorWaypoint(12, 1),
                new VisioConnectorWaypoint(12, 12), new VisioConnectorWaypoint(2.5, 1)
            })
        } };
        var document = VisioDocument.Create().GraphDiagram("Preserved", graph =>
            graph.PageSize(4, 4, unit).PreserveLayout().Import(nodes, edges));
        var page = document.Pages[0];
        var bounds = Assert.Single(page.Connectors).GetConnectorContentBounds();
        Assert.True(page.Width >= bounds.Right);
        Assert.True(page.Height >= bounds.Top);
        Assert.True(page.Width > (unit == VisioMeasurementUnit.Inches ? 12 : 12 / 2.54));
    }

    [Theory]
    [InlineData(100)]
    [InlineData(500)]
    [InlineData(1000)]
    public void RadialRingsReserveEnoughSpaceForTheirOccupancy(int count) {
        var document = VisioDocument.Create().GraphDiagram("Star", graph => {
            graph.Layout(VisioGraphLayout.Radial).Root("root", "Root");
            for (int i = 0; i < count; i++) graph.Node("n" + i, "Node " + i).Edge("root", "n" + i);
        });
        var page = document.Pages[0];
        var nodes = page.Shapes.Where(shape => shape.Id == "root" || shape.Id.StartsWith("n", StringComparison.Ordinal)).ToArray();
        Assert.Equal(count + 1, nodes.Length);
        for (int i = 0; i < nodes.Length; i++) {
            var a = nodes[i];
            Assert.InRange(a.PinX - a.Width / 2, 0, page.Width);
            Assert.InRange(a.PinY - a.Height / 2, 0, page.Height);
            Assert.InRange(a.PinX + a.Width / 2, 0, page.Width);
            Assert.InRange(a.PinY + a.Height / 2, 0, page.Height);
            for (int j = i + 1; j < nodes.Length; j++) {
                var b = nodes[j];
                Assert.True(Math.Abs(a.PinX - b.PinX) >= (a.Width + b.Width) / 2 ||
                    Math.Abs(a.PinY - b.PinY) >= (a.Height + b.Height) / 2, $"{a.Id} overlaps {b.Id}");
            }
        }
    }
}
