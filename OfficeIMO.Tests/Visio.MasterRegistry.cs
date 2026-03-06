using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioMasterRegistry {
        private static string AssetsPath => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets"));

        [Fact]
        public void RegisterMasterAllowsLookupAndShapeCreationByName() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            VisioShape blueprint = new("1", 0, 0, 2, 1, string.Empty) { NameU = "CustomProcess" };
            VisioMaster registered = document.RegisterMaster("CustomProcess", blueprint, "42");

            Assert.True(document.TryGetMaster("CustomProcess", out VisioMaster? found));
            Assert.Same(registered, found);
            Assert.Same(registered, document.GetMaster("CustomProcess"));
            Assert.Contains(document.Masters, master => ReferenceEquals(master, registered));

            VisioPage page = document.AddPage("Page-1");
            VisioShape shape = page.AddShape("shape-1", "CustomProcess", 1, 1, 2, 1, "Hello");
            document.Save();

            Assert.Same(registered, shape.Master);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape loadedShape = Assert.Single(loaded.Pages[0].Shapes);
            Assert.Equal("CustomProcess", loadedShape.Master?.NameU);
            Assert.Equal("Hello", loadedShape.Text);
        }

        [Fact]
        public void ImportMastersAndGetReturnsRegisteredMastersForReuse() {
            string template = Path.Combine(AssetsPath, "VisioTemplates", "DrawingWithShapes.vsdx");
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            IReadOnlyList<VisioMaster> imported = document.ImportMastersAndGet(template, new[] { "Rectangle", "Circle" });

            Assert.Equal(2, imported.Count);
            Assert.Contains(imported, master => master.NameU == "Rectangle");
            Assert.Contains(imported, master => master.NameU == "Circle");
            Assert.True(document.TryGetMaster("Circle", out VisioMaster? circleMaster));
            Assert.Equal("Circle", circleMaster?.NameU);

            VisioPage page = document.AddPage("Page-1");
            page.AddShape("rectangle-1", "Rectangle", 2, 2, 2, 1, "Rectangle");
            page.AddShape("circle-1", "Circle", 5, 2, 2, 2, "Circle");
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(new[] { "Rectangle", "Circle" }, loaded.Pages[0].Shapes.Select(shape => shape.Master?.NameU));
        }

        [Fact]
        public void AddShapeByMasterNameUsesPageDefaultUnit() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.RegisterMaster("MetricRect", new VisioShape("1", 0, 0, 1, 1, string.Empty) { NameU = "MetricRect" });

            VisioPage page = document.AddPage("Metric", 21, 29.7, VisioMeasurementUnit.Centimeters);
            VisioShape shape = page.AddShape("metric-1", "MetricRect", 2.54, 2.54, 5.08, 2.54, "Metric");

            Assert.Equal(5.08, shape.Width.FromInches(VisioMeasurementUnit.Centimeters), 5);
            Assert.Equal(2.54, shape.Height.FromInches(VisioMeasurementUnit.Centimeters), 5);
        }

        [Fact]
        public void AddShapeByMasterNameOnDetachedPageThrows() {
            VisioPage page = new("Detached");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                page.AddShape("shape", "Rectangle", 1, 1, 2, 1, "Detached"));

            Assert.Contains("not attached", exception.Message, StringComparison.OrdinalIgnoreCase);
        }
    }
}
