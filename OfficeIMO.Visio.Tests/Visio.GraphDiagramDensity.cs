using System;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioGraphDiagramDensityTests {
    [Fact]
    public void PreservedParallelRoutesReuseNearEndpointsWithoutLosingDistinctRoutes() {
        var nodes = new[] {
            new VisioGraphNodeRecord("a", "A") { Placement = new VisioGraphPlacement(1, 1, 1, 1) },
            new VisioGraphNodeRecord("b", "B") { Placement = new VisioGraphPlacement(3, 1, 1, 1) }
        };
        VisioGraphEdgeRecord[] edges = Enumerable.Range(0, 256).Select(index =>
            new VisioGraphEdgeRecord("route" + index, "a", "b") {
                Route = new VisioGraphRoute(new[] {
                    new VisioConnectorWaypoint(1.5, 1 + (index + 1) / 1000D),
                    new VisioConnectorWaypoint(2, 1 + (index + 1) / 1000D),
                    new VisioConnectorWaypoint(2.5, 1 + (index + 1) / 1000D)
                })
            }).Concat(new[] { new VisioGraphEdgeRecord("near-duplicate", "a", "b") {
                Route = new VisioGraphRoute(new[] {
                    new VisioConnectorWaypoint(1.5, 1.0010000005D),
                    new VisioConnectorWaypoint(2, 1.0010000005D),
                    new VisioConnectorWaypoint(2.5, 1.0010000005D)
                })
            } }).ToArray();
        var document = VisioDocument.Create().GraphDiagram("Parallel", graph =>
            graph.PageSize(5, 5).PreserveLayout().Import(nodes, edges));
        var connectors = document.Pages[0].Connectors;

        Assert.Equal(257, connectors.Count);
        Assert.Equal(256, connectors.Select(connector => connector.FromConnectionPoint).Distinct().Count());
        Assert.Equal(256, connectors.Select(connector => connector.ToConnectionPoint).Distinct().Count());
        Assert.Same(connectors[0].FromConnectionPoint, connectors[256].FromConnectionPoint);
    }

    [Fact]
    public void PreservedRouteReusesSidePointAddedAfterShapeWasIndexed() {
        var nodes = new[] {
            new VisioGraphNodeRecord("a", "A") { Placement = new VisioGraphPlacement(1, 1, 1, 1) },
            new VisioGraphNodeRecord("b", "B") { Placement = new VisioGraphPlacement(3, 1, 1, 1) },
            new VisioGraphNodeRecord("c", "C") { Placement = new VisioGraphPlacement(1, 3, 1, 1) }
        };
        var edges = new[] {
            new VisioGraphEdgeRecord("ab", "a", "b") { Route = new VisioGraphRoute(new[] {
                new VisioConnectorWaypoint(1.5, 1), new VisioConnectorWaypoint(2.5, 1)
            }) },
            new VisioGraphEdgeRecord("ca", "c", "a") { Route = new VisioGraphRoute(new[] {
                new VisioConnectorWaypoint(1, 2.5), new VisioConnectorWaypoint(1, 1.5)
            }) }
        };

        var document = VisioDocument.Create().GraphDiagram("Side points", graph =>
            graph.PageSize(5, 5).PreserveLayout().Import(nodes, edges));
        var page = document.Pages[0];
        var a = page.Shapes.Single(shape => shape.Id == "a");
        var connector = page.Connectors.Single(edge => edge.Id == "ca");
        Assert.Equal(2, a.ConnectionPoints.Count);
        Assert.Same(a.ConnectionPoints[1], connector.ToConnectionPoint);
    }

    [Theory]
    [InlineData(VisioMeasurementUnit.Centimeters)]
    [InlineData(VisioMeasurementUnit.Millimeters)]
    public void ExistingBlockCaptionRetainsPageUnitCoordinates(VisioMeasurementUnit unit) {
        var document = VisioDocument.Create().BlockDiagram("Units", graph =>
            graph.PageSize(20, 20, unit).Region("zone", "Zone caption", 0, 0, 2, 2).Block("a", "A", 0, 0));
        var page = document.Pages[0];
        var body = page.Shapes.Single(shape => shape.Id == "zone");
        var caption = page.Shapes.Single(shape => shape.Text == "Zone caption");
        Assert.Equal(body.PinX, caption.PinX, 6);
        Assert.InRange(caption.Width, 0, body.Width);
    }

    [Theory]
    [InlineData(VisioMeasurementUnit.Centimeters)]
    [InlineData(VisioMeasurementUnit.Millimeters)]
    public void PreservedGroupCaptionUsesPageUnitsOnce(VisioMeasurementUnit unit) {
        var document = VisioDocument.Create().GraphDiagram("Units", graph =>
            graph.PageSize(20, 20, unit).PreserveLayout().Import(new[] {
                new VisioGraphNodeRecord("a", "A") { Placement = new VisioGraphPlacement(10, 10, 1, 1) }
            }, Array.Empty<VisioGraphEdgeRecord>(), new[] {
                new VisioGraphClusterRecord("zone", "Zone caption", new[] { "a" }) { Placement = new VisioGraphPlacement(10, 10, 8, 8) }
            }));
        var page = document.Pages[0];
        var body = page.Shapes.Single(shape => shape.Id == "zone");
        var caption = page.Shapes.Single(shape => shape.Text == "Zone caption");
        Assert.Equal(body.PinX, caption.PinX, 6);
        Assert.InRange(caption.Width, 0, body.Width);
    }

    [Theory]
    [InlineData(VisioMeasurementUnit.Inches)]
    [InlineData(VisioMeasurementUnit.Centimeters)]
    public void PreservedPageFitTranslatesNegativeShapesAndRoutesTogether(VisioMeasurementUnit unit) {
        var document = VisioDocument.Create().GraphDiagram("Negative bounds", graph =>
            graph.PageSize(4, 4, unit).PreserveLayout().Import(new[] {
                new VisioGraphNodeRecord("a", "A") { Placement = new VisioGraphPlacement(0, 0, 1, 1) },
                new VisioGraphNodeRecord("b", "B") { Placement = new VisioGraphPlacement(3, 1, 1, 1) }
            }, new[] { new VisioGraphEdgeRecord("ab", "a", "b") {
                Label = "Negative bend", Route = new VisioGraphRoute(new[] {
                    new VisioConnectorWaypoint(.5, 0), new VisioConnectorWaypoint(-2, -2),
                    new VisioConnectorWaypoint(2.5, 1)
                })
            } }));
        var page = document.Pages[0];
        var bounds = page.GetContentBounds();
        Assert.True(bounds.Left >= 0 && bounds.Bottom >= 0);
        Assert.True(bounds.Right <= page.Width && bounds.Top <= page.Height);
        var a = page.Shapes.Single(shape => shape.Id == "a");
        var b = page.Shapes.Single(shape => shape.Id == "b");
        double scale = unit == VisioMeasurementUnit.Inches ? 1 : 1 / 2.54;
        Assert.Equal(3 * scale, b.PinX - a.PinX, 6);
        Assert.Equal(scale, b.PinY - a.PinY, 6);
        Assert.Contains(Assert.Single(page.Connectors).Waypoints, point =>
            Math.Abs(point.X - (a.PinX - 2 * scale)) < 1e-6 && Math.Abs(point.Y - (a.PinY - 2 * scale)) < 1e-6);
    }

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
