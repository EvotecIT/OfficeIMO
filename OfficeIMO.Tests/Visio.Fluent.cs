using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using SixLabors.ImageSharp;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioFluentDocumentTests {
        [Fact]
        public void CanBuildDocumentFluently() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            VisioDocument result = document.AsFluent()
                .Page("Page1", 8.5, 11, VisioMeasurementUnit.Inches, p => { })
                .End();

            Assert.Same(document, result);
            Assert.Single(document.Pages);
            Assert.Equal("Page1", document.Pages[0].Name);
            document.Save();
        }

        [Fact]
        public void FluentShapesUsePageUnitsAndPersistNumericShapeIds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            document.AsFluent()
                .Page("Metric", 10, 10, VisioMeasurementUnit.Centimeters, p => p
                    .Rect("box", 2.54, 2.54, 2.54, 2.54, "Box"))
                .End();
            Assert.Equal("box", document.Pages[0].Shapes[0].Id);
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape shape = Assert.Single(loaded.Pages[0].Shapes);
            Assert.Equal(2.54, shape.Width.FromInches(VisioMeasurementUnit.Centimeters), 5);

            using ZipArchive archive = ZipFile.OpenRead(filePath);
            using Stream pageStream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageXml = XDocument.Load(pageStream);
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            string? savedId = pageXml.Root!.Element(v + "Shapes")!.Element(v + "Shape")!.Attribute("ID")?.Value;
            Assert.True(int.TryParse(savedId, out _));
        }

        [Fact]
        public void FluentConnectCanTargetSidesAndStyleLines() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            document.AsFluent()
                .Page("Page1", p => p
                    .Rect("left", 1, 1, 2, 1, "Left")
                    .Rect("right", 5, 1, 2, 1, "Right")
                    .Connect("left", "right", VisioSide.Right, VisioSide.Left, c => c
                        .RightAngle()
                        .LineColor(Color.DarkBlue)
                        .ArrowEnd(EndArrow.Triangle)))
                .End();
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector connector = Assert.Single(loaded.Pages[0].Connectors);
            Assert.Equal(ConnectorKind.RightAngle, connector.Kind);
            Assert.NotNull(connector.FromConnectionPoint);
            Assert.NotNull(connector.ToConnectionPoint);
            Assert.Equal(Color.DarkBlue, connector.LineColor);
            Assert.Equal(EndArrow.Triangle, connector.EndArrow);
        }

        [Fact]
        public void SideSelectionUsesNamedSidePointEvenWhenCustomPointsExist() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            document.AsFluent()
                .Page("Page1", p => p
                    .Rect("left", 1, 1, 2, 2, "Left")
                    .Rect("right", 5, 1, 2, 2, "Right"))
                .End();

            VisioShape left = document.Pages[0].Shapes[0];
            VisioShape right = document.Pages[0].Shapes[1];
            left.ConnectionPoints.Add(new VisioConnectionPoint(0.25, 0.25, 0, 0));
            right.ConnectionPoints.Add(new VisioConnectionPoint(1.75, 1.75, 0, 0));

            VisioConnector connector = document.Pages[0].AddConnector(left, right, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            Assert.Equal(left.Width, connector.FromConnectionPoint!.X, 5);
            Assert.Equal(left.Height / 2, connector.FromConnectionPoint.Y, 5);
            Assert.Equal(0, connector.ToConnectionPoint!.X, 5);
            Assert.Equal(right.Height / 2, connector.ToConnectionPoint.Y, 5);
        }
    }
}
