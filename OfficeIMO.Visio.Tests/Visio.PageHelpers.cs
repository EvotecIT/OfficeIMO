using System;
using System.IO;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioPageHelpers {
        [Fact]
        public void AddShapeAndConnectorPopulateCollections() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1", OfficeLength.Inches(8.5), OfficeLength.Inches(11));
            VisioMaster master = new("1", "Rectangle", new VisioShape("1"));

            VisioShape shape1 = page.AddShape("1", master, OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Start");
            VisioShape shape2 = page.AddShape("2", master, OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "End");

            Assert.Equal(8.5, page.Width.ToInches());
            Assert.Equal(11, page.Height.ToInches());
            Assert.Equal(2, page.Shapes.Count);
            Assert.Contains(shape1, page.Shapes);
            Assert.Contains(shape2, page.Shapes);

            VisioConnector connector = page.AddConnector("3", shape1, shape2, ConnectorKind.Straight);

            Assert.Single(page.Connectors);
            Assert.Contains(connector, page.Connectors);
            Assert.Equal(ConnectorKind.Straight, connector.Kind);
            document.Save();
        }

        [Fact]
        public void SizeAndGridHelpersSetProperties() {
            VisioPage page = new("Page-1");
            page.Size(OfficeLength.Inches(10), OfficeLength.Inches(5)).Grid(true, false);
            Assert.Equal(10, page.Width.ToInches());
            Assert.Equal(5, page.Height.ToInches());
            Assert.True(page.GridVisible);
            Assert.False(page.Snap);
        }

        [Fact]
        public void MetricUnitsConvertToInches() {
            VisioDocument document = VisioDocument.Create(Path.Combine(
                Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Metric", OfficeLength.Centimeters(21), OfficeLength.Centimeters(29.7), scaleUnit: VisioMeasurementUnit.Centimeters);
            Assert.Equal(21, Math.Round(page.Width.ToCentimeters(), 2));
            Assert.Equal(29.7, Math.Round(page.Height.ToCentimeters(), 2));
            Assert.Equal(8.27, Math.Round(page.Width.ToInches(), 2));
            Assert.Equal(11.69, Math.Round(page.Height.ToInches(), 2));
        }
    }
}
