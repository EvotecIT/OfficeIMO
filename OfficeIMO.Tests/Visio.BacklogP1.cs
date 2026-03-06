using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioBacklogP1 {
        [Fact]
        public void DefaultConnectorKindIsDynamicAcrossApis() {
            VisioPage page = new("Page-1");
            VisioShape left = new("1", 1, 1, 1, 1, "Left");
            VisioShape right = new("2", 3, 1, 1, 1, "Right");
            page.Shapes.Add(left);
            page.Shapes.Add(right);

            VisioConnector constructorConnector = new(left, right);
            VisioConnector pageConnector = page.AddConnector(left, right);

            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            document.AsFluent()
                .Page("Page1", p => p
                    .Rect("left", 1, 1, 1, 1, "Left")
                    .Rect("right", 3, 1, 1, 1, "Right")
                    .Connect("left", "right"))
                .End();

            VisioConnector fluentConnector = Assert.Single(document.Pages[0].Connectors);

            Assert.Equal(ConnectorKind.Dynamic, constructorConnector.Kind);
            Assert.Equal(ConnectorKind.Dynamic, pageConnector.Kind);
            Assert.Equal(ConnectorKind.Dynamic, fluentConnector.Kind);
        }

        [Fact]
        public void SaveDoesNotMutateAutoMasteredShapeState() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.UseMastersByDefault = true;

            VisioPage page = document.AddPage("Page-1");
            VisioShape shape = new("shape") {
                NameU = "Rectangle",
                PinX = 2,
                PinY = 3
            };
            page.Shapes.Add(shape);

            document.Save();

            Assert.Null(shape.Master);
            Assert.Equal(0, shape.Width);
            Assert.Equal(0, shape.Height);
            Assert.Equal(0, shape.LocPinX);
            Assert.Equal(0, shape.LocPinY);
        }

        [Fact]
        public void FluentDuplicateShapeIdsThrowHelpfulError() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.AsFluent()
                    .Page("Page1", p => p
                        .Rect("dup", 1, 1, 1, 1, "Left")
                        .Circle("dup", 3, 1, 1, "Right"))
                    .End());

            Assert.Contains("dup", exception.Message);
        }

        [Fact]
        public void SaveRejectsDuplicateIdsAcrossPageObjects() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");

            page.Shapes.Add(new VisioShape("dup", 1, 1, 1, 1, "First"));
            page.Shapes.Add(new VisioShape("dup", 3, 1, 1, 1, "Second"));

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => document.Save());
            Assert.Contains("dup", exception.Message);
        }

        [Fact]
        public void MastersWithCollidingIdsAreRemappedAndRoundTrip() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");

            VisioMaster rectangleMaster = new("2", "Rectangle", new VisioShape("1", 0, 0, 2, 1, string.Empty) { NameU = "Rectangle" });
            VisioMaster ellipseMaster = new("2", "Ellipse", new VisioShape("1", 0, 0, 2, 1, string.Empty) { NameU = "Ellipse" });

            VisioShape rectangle = new("1", 1, 1, 2, 1, "Rectangle") { NameU = "Rectangle", Master = rectangleMaster };
            VisioShape ellipse = new("2", 4, 1, 2, 1, "Ellipse") { NameU = "Ellipse", Master = ellipseMaster };
            page.Shapes.Add(rectangle);
            page.Shapes.Add(ellipse);

            document.Save();

            using (Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read)) {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XDocument pageDoc = XDocument.Load(package.GetPart(new Uri("/visio/pages/page1.xml", UriKind.Relative)).GetStream());
                XElement[] shapes = pageDoc.Root?.Element(ns + "Shapes")?.Elements(ns + "Shape").ToArray() ?? Array.Empty<XElement>();
                string? firstMasterId = shapes[0].Attribute("Master")?.Value;
                string? secondMasterId = shapes[1].Attribute("Master")?.Value;

                Assert.False(string.IsNullOrWhiteSpace(firstMasterId));
                Assert.False(string.IsNullOrWhiteSpace(secondMasterId));
                Assert.NotEqual(firstMasterId, secondMasterId);

                XDocument mastersDoc = XDocument.Load(package.GetPart(new Uri("/visio/masters/masters.xml", UriKind.Relative)).GetStream());
                string[] masterIds = mastersDoc.Root?.Elements(ns + "Master").Select(m => m.Attribute("ID")?.Value ?? string.Empty).ToArray() ?? Array.Empty<string>();
                Assert.Equal(2, masterIds.Length);
                Assert.Equal(2, masterIds.Distinct(StringComparer.Ordinal).Count());
            }

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal("Rectangle", loaded.Pages[0].Shapes[0].Master?.NameU);
            Assert.Equal("Ellipse", loaded.Pages[0].Shapes[1].Master?.NameU);
        }
    }
}
