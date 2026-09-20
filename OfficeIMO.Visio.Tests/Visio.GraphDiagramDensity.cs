using System;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioGraphDiagramDensityTests {
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
