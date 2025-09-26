using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioMultiPage {
        [Fact]
        public void CanRoundTripMultiplePages() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            try {
                VisioDocument document = VisioDocument.Create(filePath);

                VisioPage firstPage = document.AddPage("Page1");
                firstPage.ViewScale = 1.25;
                firstPage.ViewCenterX = 4;
                firstPage.ViewCenterY = 5;

                VisioShape firstShape = new("1", 1, 1, 2, 1, "First");
                firstShape.Name = "FirstShape";
                VisioShape secondShape = new("2", 4, 3, 1.5, 1, "Second");
                secondShape.Name = "SecondShape";
                firstPage.Shapes.Add(firstShape);
                firstPage.Shapes.Add(secondShape);

                VisioConnector connector = new("C1", firstShape, secondShape) {
                    Kind = ConnectorKind.RightAngle,
                    Label = "Connector 1"
                };
                firstPage.Connectors.Add(connector);

                VisioPage secondPage = document.AddPage("Page2");
                secondPage.ViewScale = 0.75;
                secondPage.ViewCenterX = 3;
                secondPage.ViewCenterY = 2.5;

                VisioShape thirdShape = new("3", 2, 2, 1, 1, "Third");
                VisioShape fourthShape = new("4", 5, 1.5, 1.25, 1.25, "Fourth");
                secondPage.Shapes.Add(thirdShape);
                secondPage.Shapes.Add(fourthShape);

                document.Save();

                VisioDocument reloaded = VisioDocument.Load(filePath);

                Assert.Equal(2, reloaded.Pages.Count);
                Assert.Equal(new[] { "Page1", "Page2" }, reloaded.Pages.Select(p => p.Name));

                VisioPage reloadedFirstPage = reloaded.Pages[0];
                Assert.Equal(1.25, reloadedFirstPage.ViewScale, 5);
                Assert.Equal(4, reloadedFirstPage.ViewCenterX, 5);
                Assert.Equal(5, reloadedFirstPage.ViewCenterY, 5);
                Assert.Equal(new[] { "1", "2" }, reloadedFirstPage.Shapes.Select(s => s.Id));
                Assert.Single(reloadedFirstPage.Connectors);
                VisioConnector reloadedConnector = reloadedFirstPage.Connectors.Single();
                Assert.Equal("C1", reloadedConnector.Id);
                Assert.Equal(ConnectorKind.RightAngle, reloadedConnector.Kind);
                Assert.Equal("Connector 1", reloadedConnector.Label);
                Assert.Equal("1", reloadedConnector.From.Id);
                Assert.Equal("2", reloadedConnector.To.Id);

                VisioPage reloadedSecondPage = reloaded.Pages[1];
                Assert.Equal(0.75, reloadedSecondPage.ViewScale, 5);
                Assert.Equal(3, reloadedSecondPage.ViewCenterX, 5);
                Assert.Equal(2.5, reloadedSecondPage.ViewCenterY, 5);
                Assert.Equal(new[] { "3", "4" }, reloadedSecondPage.Shapes.Select(s => s.Id));
                Assert.Empty(reloadedSecondPage.Connectors);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ContentTypesIncludesAllPageOverrides() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            try {
                VisioDocument document = VisioDocument.Create(filePath);
                document.AddPage("Alpha");
                document.AddPage("Beta");
                document.AddPage("Gamma");

                document.Save();

                using FileStream stream = File.OpenRead(filePath);
                using ZipArchive archive = new(stream, ZipArchiveMode.Read);
                ZipArchiveEntry? contentTypesEntry = archive.GetEntry("[Content_Types].xml");
                Assert.NotNull(contentTypesEntry);

                List<string> pagePartNames = archive.Entries
                    .Where(entry => entry.FullName.StartsWith("visio/pages/", StringComparison.OrdinalIgnoreCase)
                        && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(entry.FullName, "visio/pages/pages.xml", StringComparison.OrdinalIgnoreCase))
                    .Select(entry => "/" + entry.FullName.TrimStart('/'))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                Assert.Equal(3, pagePartNames.Count);

                using Stream contentStream = contentTypesEntry!.Open();
                XDocument documentXml = XDocument.Load(contentStream);
                XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
                List<XElement> overrideElements = documentXml.Root?
                    .Elements(ct + "Override")
                    .ToList() ?? new List<XElement>();

                foreach (string partName in pagePartNames) {
                    List<XElement> matches = overrideElements
                        .Where(element => string.Equals((string?)element.Attribute("PartName"), partName, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    Assert.Single(matches);
                    Assert.Equal("application/vnd.ms-visio.page+xml", (string?)matches[0].Attribute("ContentType"));
                }

                int uniquePageOverrideCount = overrideElements
                    .Where(element => string.Equals((string?)element.Attribute("ContentType"), "application/vnd.ms-visio.page+xml", StringComparison.OrdinalIgnoreCase))
                    .Select(element => (string?)element.Attribute("PartName"))
                    .Where(name => !string.IsNullOrEmpty(name))
                    .Select(name => name!)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count();

                Assert.Equal(pagePartNames.Count, uniquePageOverrideCount);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
