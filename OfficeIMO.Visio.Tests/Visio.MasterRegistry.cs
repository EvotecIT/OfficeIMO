using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioMasterRegistry {
        private static string AssetsPath => Path.Combine(RepositoryTestPaths.Find(), "Assets");

        [Fact]
        public void RegisterMasterAllowsLookupAndShapeCreationByName() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            VisioShape blueprint = new("1", OfficeLength.Inches(0), OfficeLength.Inches(0), OfficeLength.Inches(2), OfficeLength.Inches(1), string.Empty);
            VisioMaster registered = document.RegisterMaster("CustomProcess", blueprint, "42");

            Assert.True(document.TryGetMaster("CustomProcess", out VisioMaster? found));
            Assert.Same(registered, found);
            Assert.Same(registered, document.GetMaster("CustomProcess"));
            Assert.Contains(document.Masters, master => ReferenceEquals(master, registered));

            VisioPage page = document.AddPage("Page-1");
            VisioShape shape = page.AddShape("shape-1", "CustomProcess", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Hello");
            document.Save();

            Assert.Same(registered, shape.Master);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape loadedShape = Assert.Single(loaded.Pages[0].Shapes);
            Assert.Equal("CustomProcess", loadedShape.Master?.NameU);
            Assert.Equal("Hello", loadedShape.Text);
        }

        [Fact]
        public void RegisterMasterPreservesDistinctMastersWithDuplicateNameU() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioMaster first = new("1", "DuplicateName", new VisioShape("1", OfficeLength.Inches(0), OfficeLength.Inches(0), OfficeLength.Inches(1), OfficeLength.Inches(1), string.Empty));
            VisioMaster second = new("2", "DuplicateName", new VisioShape("1", OfficeLength.Inches(0), OfficeLength.Inches(0), OfficeLength.Inches(2), OfficeLength.Inches(1), string.Empty));

            document.RegisterMaster(first);
            document.RegisterMaster(second);

            Assert.Equal(2, document.Masters.Count(master => master.NameU == "DuplicateName"));
            Assert.True(document.TryGetMaster("DuplicateName", out VisioMaster? found));
            Assert.Same(second, found);
        }

        [Fact]
        public void RegisterMasterPreservesDistinctMastersWithDuplicatePackageIds() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioMaster first = new("1", "PackageA", new VisioShape("1", OfficeLength.Inches(0), OfficeLength.Inches(0), OfficeLength.Inches(1), OfficeLength.Inches(1), string.Empty));
            VisioMaster second = new("1", "PackageB", new VisioShape("1", OfficeLength.Inches(0), OfficeLength.Inches(0), OfficeLength.Inches(2), OfficeLength.Inches(1), string.Empty));

            document.RegisterMaster(first);
            document.RegisterMaster(second);

            Assert.Contains(document.Masters, master => ReferenceEquals(master, first));
            Assert.Contains(document.Masters, master => ReferenceEquals(master, second));
            Assert.Equal(2, document.Masters.Count(master => master.Id == "1"));
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
            page.AddShape("rectangle-1", "Rectangle", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Rectangle");
            page.AddShape("circle-1", "Circle", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), "Circle");
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(new[] { "Rectangle", "Circle" }, loaded.Pages[0].Shapes.Select(shape => shape.Master?.NameU));
        }

        [Fact]
        public void AddShapeByMasterNameUsesPageDefaultUnit() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.RegisterMaster("MetricRect", new VisioShape("1", OfficeLength.Inches(0), OfficeLength.Inches(0), OfficeLength.Inches(1), OfficeLength.Inches(1), string.Empty));

            VisioPage page = document.AddPage("Metric", OfficeLength.Centimeters(21), OfficeLength.Centimeters(29.7), scaleUnit: VisioMeasurementUnit.Centimeters);
            VisioShape shape = page.AddShape("metric-1", "MetricRect", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.54, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.54, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5.08, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.54, page.DefaultUnit)), "Metric");

            Assert.Equal(5.08, shape.Width.ToInches().FromInches(VisioMeasurementUnit.Centimeters), 5);
            Assert.Equal(2.54, shape.Height.ToInches().FromInches(VisioMeasurementUnit.Centimeters), 5);
        }

        [Fact]
        public void AddShapeByMasterNameOnDetachedPageThrows() {
            VisioPage page = new("Detached");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                page.AddShape("shape", "Rectangle", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Detached"));

            Assert.Contains("not attached", exception.Message, StringComparison.OrdinalIgnoreCase);
        }
    }
}
