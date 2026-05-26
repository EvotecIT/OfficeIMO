using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests {
    public class VisioCalloutTests {
        [Fact]
        public void CalloutsCreateSemanticShapeLeaderLayerAndRoundTrip() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string roundTripPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Annotations", 11, 8.5);
            VisioShape service = page.AddProcess(4, 4.5, 2, 1, "API");

            VisioShape callout = page.AddCallout(service, "api-note", "Check retry policy", 7.5, 6, new VisioCalloutOptions {
                Width = 2.4,
                Height = 0.8,
                ShapeStyle = new VisioShapeStyle(Color.LightYellow, Color.DarkOrange, 0.02),
                LeaderStyle = new VisioConnectorStyle(Color.DarkOrange, 0.015, 2, EndArrow.None),
                RouteOffset = 0.15
            });

            Assert.True(callout.IsCallout);
            Assert.Equal(service.Id, callout.CalloutTargetId);
            Assert.Equal("Callout", callout.GetUserCellValue("OfficeIMO.Kind"));
            Assert.Equal("api-note", page.Callouts().Single().Id);
            Assert.Single(page.SelectCallouts());
            Assert.Contains("Annotations", callout.LayerNames);

            VisioConnector leader = page.Connectors.Single();
            Assert.Same(callout, leader.From);
            Assert.Same(service, leader.To);
            Assert.Contains("Annotations", leader.LayerNames);
            Assert.Equal(ConnectorKind.RightAngle, leader.Kind);
            Assert.Equal(EndArrow.None, leader.EndArrow);
            Assert.Equal(2, leader.Waypoints.Count);
            Assert.Equal(leader.Id, callout.GetUserCellValue("OfficeIMO.CalloutLeaderId"));

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertCalloutXml(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages.Single();
            VisioShape loadedCallout = loadedPage.Callouts().Single();
            Assert.True(loadedCallout.IsCallout);
            Assert.Equal("Check retry policy", loadedCallout.Text);
            Assert.Equal(loadedPage.Shapes.Single(shape => shape.Text == "API").Id, loadedCallout.CalloutTargetId);
            Assert.Single(loadedPage.Connectors);
            Assert.Equal(2, loadedPage.Connectors.Single().Waypoints.Count);

            loaded.Save(roundTripPath);
            Assert.Empty(VisioValidator.Validate(roundTripPath));
            AssertCalloutXml(roundTripPath);
        }

        [Fact]
        public void CalloutsCanUseAutomaticIdsAndStraightLeadersWithoutDefaultLayer() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Auto");
            VisioShape target = page.AddRectangle(2, 2, 1.5, 0.8, "Target");

            VisioShape callout = page.AddCallout(target, "Plain note", 4.5, 2.5, new VisioCalloutOptions {
                LeaderKind = ConnectorKind.Straight,
                RouteLeader = false,
                LayerName = null
            });

            Assert.True(callout.IsCallout);
            Assert.NotEqual(target.Id, callout.Id);
            Assert.Empty(callout.LayerNames);
            VisioConnector leader = page.Connectors.Single();
            Assert.Equal(ConnectorKind.Straight, leader.Kind);
            Assert.Empty(leader.Waypoints);
            Assert.Empty(leader.LayerNames);
        }

        private static void AssertCalloutXml(string filePath) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument page = ReadXml(archive, "visio/pages/page1.xml");

            XElement callout = page.Descendants(ns + "Shape")
                .Single(shape => shape.Element(ns + "Text")?.Value == "Check retry policy");
            XElement userSection = callout.Elements(ns + "Section")
                .Single(section => (string?)section.Attribute("N") == "User");

            Assert.Equal("Callout", UserCellValue(userSection, ns, "OfficeIMO.Kind"));
            Assert.False(string.IsNullOrWhiteSpace(UserCellValue(userSection, ns, "OfficeIMO.CalloutTargetId")));
            Assert.False(string.IsNullOrWhiteSpace(UserCellValue(userSection, ns, "OfficeIMO.CalloutLeaderId")));

            XDocument pages = ReadXml(archive, "visio/pages/pages.xml");
            XElement layerSection = pages.Descendants(ns + "Section")
                .Single(section => (string?)section.Attribute("N") == "Layer");
            Assert.Contains(layerSection.Elements(ns + "Row"), row =>
                row.Elements(ns + "Cell").Any(cell =>
                    (string?)cell.Attribute("N") == "Name" &&
                    (string?)cell.Attribute("V") == "Annotations"));
        }

        private static string UserCellValue(XElement userSection, XNamespace ns, string rowName) {
            XElement row = userSection.Elements(ns + "Row")
                .Single(element => (string?)element.Attribute("N") == rowName);
            return row.Elements(ns + "Cell")
                .Single(cell => (string?)cell.Attribute("N") == "Value")
                .Attribute("V")!
                .Value;
        }

        private static XDocument ReadXml(ZipArchive archive, string entryName) {
            ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Missing " + entryName);
            using Stream stream = entry.Open();
            return XDocument.Load(stream);
        }
    }
}
