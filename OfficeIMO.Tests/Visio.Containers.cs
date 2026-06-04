using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests {
    public class VisioContainerTests {
        [Fact]
        public void ContainersSaveNativeUserCellsRelationshipsAndLoadAsSemanticShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string roundTripPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Containers", 11, 8.5);
            VisioShape api = page.AddRectangle(3, 5.5, 1.7, 0.8, "API");
            VisioShape worker = page.AddRectangle(6, 5.5, 1.7, 0.8, "Worker");

            VisioShape container = page.AddContainer("app-tier", "Application tier", new[] { api, worker }, new VisioContainerOptions {
                Margin = 0.3,
                HeadingHeight = 0.4,
                FillColor = Color.LightCyan,
                LineColor = Color.DodgerBlue
            });
            container.SetUserCell("OfficeIMO.Role", "Tier", "STR", prompt: "OfficeIMO semantic role");
            page.SelectContainers().Stroke(Color.DodgerBlue, 0.02);
            page.SelectWithUserCell("OfficeIMO.Role", "Tier").UserCell("OfficeIMO.Reviewed", "Yes", "STR");

            Assert.True(container.IsContainer);
            Assert.Equal(new[] { api.Id, worker.Id }, container.ContainerMemberIds.ToArray());
            Assert.Contains(container.Id, api.ContainerOwnerIds);
            Assert.Single(page.Containers());
            Assert.Single(page.ShapesWithUserCell("OfficeIMO.Role", "Tier"));

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertContainerXml(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages[0];
            VisioShape loadedContainer = Assert.Single(loadedPage.Containers());
            Assert.True(loadedContainer.IsContainer);
            Assert.Equal("Container", loadedContainer.GetUserCellValue("msvStructureType"));
            Assert.Equal("Yes", loadedContainer.GetUserCellValue("OfficeIMO.Reviewed"));
            Assert.Equal(2, loadedContainer.ContainerMemberIds.Count);
            Assert.All(loadedContainer.ContainerMemberIds, memberId => Assert.Contains(loadedContainer.Id, loadedPage.FindShapeById(memberId)!.ContainerOwnerIds));

            loaded.Save(roundTripPath);
            Assert.Empty(VisioValidator.Validate(roundTripPath));
            AssertContainerXml(roundTripPath);
        }

        [Fact]
        public void ContainersConvertPageUnitOptionsToInches() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Metric Containers", 20, 15, VisioMeasurementUnit.Centimeters);
            VisioShape api = page.AddRectangle(5, 5, 2.54, 2.54, "API");
            VisioShape worker = page.AddRectangle(10, 5, 2.54, 2.54, "Worker");
            OfficeIMO.Visio.VisioShapeBounds memberBounds = new[] { api, worker }.GetShapeBounds();

            VisioShape container = page.AddContainer("metric-tier", "Metric tier", new[] { api, worker }, new VisioContainerOptions {
                Margin = 1.27,
                HeadingHeight = 2.54
            });

            Assert.Equal("0.5", container.GetUserCellValue("msvSDContainerMargin"));
            Assert.Equal(memberBounds.Width + 1.0, container.Width, 6);
            Assert.Equal(memberBounds.Height + 2.0, container.Height, 6);

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertContainerMarginXml(filePath, "Metric tier", "0.5");
        }

        [Fact]
        public void LoadedContainersCanEditMembershipRefitAndSaveNativeRelationships() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Loaded container editing", 11, 8.5);
            VisioShape api = AddRect(page, "api", 2, 5, 1.4, 0.7, "API");
            VisioShape worker = AddRect(page, "worker", 4.5, 5, 1.4, 0.7, "Worker");
            AddRect(page, "db", 7, 4.2, 1.4, 0.7, "Database");
            page.AddContainer("tier", "Application tier", new[] { api, worker });
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages[0];
            VisioShape loadedContainer = loadedPage.FindShapeById("tier")!;
            VisioShape loadedApi = loadedPage.FindShapeById("api")!;
            VisioShape loadedDb = loadedPage.FindShapeById("db")!;

            loadedPage.AddToContainer(loadedContainer, loadedDb, resizeToFit: true, resizeOptions: new VisioContainerOptions {
                Margin = 0.4,
                HeadingHeight = 0.25
            });
            loadedPage.RemoveFromContainer(loadedContainer, loadedApi, resizeToFit: true, resizeOptions: new VisioContainerOptions {
                Margin = 0.4,
                HeadingHeight = 0.25
            });
            loaded.Save(updatedPath);

            Assert.Empty(VisioValidator.Validate(updatedPath));
            AssertContainerMembershipXml(updatedPath, "Application tier", new[] { "Worker", "Database" }, new[] { "API" });

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage updatedPage = updated.Pages[0];
            VisioShape updatedContainer = updatedPage.FindShapeById("tier")!;
            Assert.Equal(new[] { "db", "worker" }, updatedPage.GetContainerMembers(updatedContainer).Select(shape => shape.Id).OrderBy(id => id, StringComparer.Ordinal).ToArray());
            Assert.DoesNotContain("tier", updatedPage.FindShapeById("api")!.ContainerOwnerIds);
            Assert.Contains("tier", updatedPage.FindShapeById("db")!.ContainerOwnerIds);
            Assert.True(updatedContainer.Width > 4D);
        }

        [Fact]
        public void LoadedContainerMetadataStyleAndHeadingCanBeUpdatedAndRoundTrip() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Container metadata", 11, 8.5);
            VisioShape api = AddRect(page, "api", 2, 5, 1.4, 0.7, "API");
            VisioShape worker = AddRect(page, "worker", 4.5, 5, 1.4, 0.7, "Worker");
            page.AddContainer("tier", "Runtime tier", new[] { api, worker }, new VisioContainerOptions {
                Margin = 0.2,
                HeadingHeight = 0.3
            });
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages[0];
            VisioShape loadedContainer = loadedPage.FindShapeById("tier")!;
            VisioContainerInfo before = loadedPage.GetContainerInfo(loadedContainer);

            Assert.Equal(2, before.MemberCount);
            Assert.Equal(0.2, before.Margin, 6);
            Assert.Equal(0.3, before.HeadingHeight, 6);

            loadedPage.ConfigureContainer(loadedContainer, options => {
                options.Margin = 0.45;
                options.HeadingHeight = 0.55;
                options.AutoResize = false;
                options.Locked = true;
                options.NoHighlight = true;
                options.NoRibbon = true;
                options.ContainerStyle = 7;
                options.HeadingStyle = 3;
                options.ShapeStyle = new VisioShapeStyle(Color.LightYellow, Color.DarkBlue, 0.03, linePattern: 2, fillPattern: 1) {
                    TextStyle = new VisioTextStyle {
                        Color = Color.DarkBlue,
                        Size = 13,
                        Bold = true
                    }
                };
            }, refit: true);
            loaded.Save(updatedPath);

            Assert.Empty(VisioValidator.Validate(updatedPath));
            AssertContainerMetadataXml(updatedPath, "Runtime tier", "0.45", "0.55", "0", "1", "1", "1", "7", "3");

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage updatedPage = updated.Pages[0];
            VisioShape updatedContainer = updatedPage.FindShapeById("tier")!;
            VisioContainerInfo info = updatedPage.GetContainerInfo(updatedContainer);

            Assert.Equal(2, info.MemberCount);
            Assert.Equal(0.45, info.Margin, 6);
            Assert.Equal(0.55, info.HeadingHeight, 6);
            Assert.False(info.AutoResize);
            Assert.True(info.Locked);
            Assert.True(info.NoHighlight);
            Assert.True(info.NoRibbon);
            Assert.Equal(7, info.ContainerStyle);
            Assert.Equal(3, info.HeadingStyle);
            Assert.Equal(Color.LightYellow, updatedContainer.FillColor);
            Assert.Equal(Color.DarkBlue, updatedContainer.LineColor);
            Assert.Equal(0.03, updatedContainer.LineWeight, 6);
            Assert.Equal(2, updatedContainer.LinePattern);
            Assert.NotNull(updatedContainer.TextStyle);
            Assert.Equal(13, updatedContainer.TextStyle!.Size);
            Assert.Equal(true, updatedContainer.TextStyle.Bold);
        }

        [Fact]
        public void LoadedMetricContainerRefitKeepsStoredMarginAndHeadingInInches() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Metric refit", 20, 15, VisioMeasurementUnit.Centimeters);
            VisioShape api = page.AddRectangle(5, 5, 2.54, 2.54, "API");
            VisioShape worker = page.AddRectangle(10, 5, 2.54, 2.54, "Worker");
            page.AddContainer("metric-tier", "Metric tier", new[] { api, worker }, new VisioContainerOptions {
                Margin = 1.27,
                HeadingHeight = 2.54
            });
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages[0];
            VisioShape loadedContainer = loadedPage.FindShapeById("metric-tier")!;
            OfficeIMO.Visio.VisioShapeBounds memberBounds = loadedPage.GetContainerMembers(loadedContainer).GetShapeBounds();

            loadedPage.RefitContainer(loadedContainer);

            Assert.Equal(memberBounds.Width + 1.0, loadedContainer.Width, 6);
            Assert.Equal(memberBounds.Height + 2.0, loadedContainer.Height, 6);
            Assert.Equal(1.27, loadedPage.GetContainerInfo(loadedContainer).Margin, 6);
            Assert.Equal(2.54, loadedPage.GetContainerInfo(loadedContainer).HeadingHeight, 6);
        }

        [Fact]
        public void FluentContainerEditingKeepsLoadedContainersEasyToMaintainById() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Fluent containers", page => page
                    .Rect("api", 2, 4, 1.3, 0.7, "API")
                    .Rect("db", 4, 4, 1.3, 0.7, "Database")
                    .Container("tier", "Runtime tier", new[] { "api", "db" }, options => {
                        options.Margin = 0.25;
                        options.HeadingHeight = 0.35;
                    }))
                .End()
                .Save();

            VisioDocument.Load(filePath)
                .AsFluent()
                .ExistingPage("Fluent containers", page => page
                    .Rect("cache", 6, 4, 1.3, 0.7, "Cache")
                    .AddToContainer("tier", new[] { "cache" }, options => {
                        options.Margin = 0.35;
                        options.HeadingHeight = 0.3;
                    })
                    .RemoveFromContainer("tier", new[] { "db" }, options => {
                        options.Margin = 0.35;
                        options.HeadingHeight = 0.3;
                    }))
                .End()
                .Save(updatedPath);

            Assert.Empty(VisioValidator.Validate(updatedPath));
            AssertContainerMembershipXml(updatedPath, "Runtime tier", new[] { "API", "Cache" }, new[] { "Database" });

            VisioPage page = VisioDocument.Load(updatedPath).Pages[0];
            VisioShape container = page.FindShapeById("tier")!;
            Assert.Equal(new[] { "api", "cache" }, page.GetContainerMembers(container).Select(shape => shape.Id).OrderBy(id => id, StringComparer.Ordinal).ToArray());
            Assert.DoesNotContain("tier", page.FindShapeById("db")!.ContainerOwnerIds);
            Assert.Contains("tier", page.FindShapeById("cache")!.ContainerOwnerIds);
        }

        [Fact]
        public void FluentCanInspectAndConfigureLoadedContainerMetadataById() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            int seenMemberCount = 0;

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Fluent metadata", page => page
                    .Rect("api", 2, 4, 1.3, 0.7, "API")
                    .Rect("db", 4, 4, 1.3, 0.7, "Database")
                    .Container("tier", "Runtime tier", new[] { "api", "db" }))
                .End()
                .Save();

            VisioDocument.Load(filePath)
                .AsFluent()
                .ExistingPage("Fluent metadata", page => {
                    seenMemberCount = page.ContainerInfo("tier").MemberCount;
                    page.ConfigureContainer("tier", options => {
                        options.Margin = 0.4;
                        options.HeadingHeight = 0.25;
                        options.NoRibbon = true;
                        options.ShapeStyle = new VisioShapeStyle(Color.LightCyan, Color.DodgerBlue, 0.025);
                    }, refit: true);
                })
                .End()
                .Save(updatedPath);

            Assert.Equal(2, seenMemberCount);
            Assert.Empty(VisioValidator.Validate(updatedPath));
            AssertContainerMetadataXml(updatedPath, "Runtime tier", "0.4", "0.25", "1", "0", "0", "1", "1", "1");

            VisioShape container = VisioDocument.Load(updatedPath).Pages[0].FindShapeById("tier")!;
            Assert.Equal(Color.LightCyan, container.FillColor);
            Assert.Equal(Color.DodgerBlue, container.LineColor);
            Assert.Equal(0.025, container.LineWeight, 6);
        }

        private static VisioShape AddRect(VisioPage page, string id, double x, double y, double width, double height, string text) {
            VisioShape shape = new(id, x, y, width, height, text) {
                NameU = "Rectangle"
            };
            page.Shapes.Add(shape);
            return shape;
        }

        private static void AssertContainerXml(string filePath) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument page = ReadXml(archive, "visio/pages/page1.xml");

            XElement container = page.Descendants(ns + "Shape")
                .Single(shape => shape.Element(ns + "Text")?.Value == "Application tier");
            string containerId = container.Attribute("ID")!.Value;

            XElement userSection = container.Elements(ns + "Section")
                .Single(section => (string?)section.Attribute("N") == "User");
            Assert.Equal("Container", UserCellValue(userSection, ns, "msvStructureType"));
            Assert.Equal("0.3", UserCellValue(userSection, ns, "msvSDContainerMargin"));
            Assert.Equal("1", UserCellValue(userSection, ns, "msvSDContainerResize"));
            Assert.Equal("Tier", UserCellValue(userSection, ns, "OfficeIMO.Role"));
            Assert.Equal("Yes", UserCellValue(userSection, ns, "OfficeIMO.Reviewed"));

            XElement relationshipCell = container.Elements(ns + "Cell")
                .Single(cell => (string?)cell.Attribute("N") == "Relationships");
            string formula = relationshipCell.Attribute("F")!.Value;
            Assert.Contains("DEPENDSON(1,Sheet.", formula);

            XElement[] memberShapes = page.Descendants(ns + "Shape")
                .Where(shape => shape.Element(ns + "Text")?.Value is "API" or "Worker")
                .ToArray();
            Assert.Equal(2, memberShapes.Length);
            foreach (XElement memberShape in memberShapes) {
                XElement memberRelationship = memberShape.Elements(ns + "Cell")
                    .Single(cell => (string?)cell.Attribute("N") == "Relationships");
                Assert.Contains($"DEPENDSON(4,Sheet.{containerId}!SheetRef())", memberRelationship.Attribute("F")!.Value);
            }
        }

        private static void AssertContainerMarginXml(string filePath, string text, string expectedMargin) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument page = ReadXml(archive, "visio/pages/page1.xml");

            XElement container = page.Descendants(ns + "Shape")
                .Single(shape => shape.Element(ns + "Text")?.Value == text);
            XElement userSection = container.Elements(ns + "Section")
                .Single(section => (string?)section.Attribute("N") == "User");
            Assert.Equal(expectedMargin, UserCellValue(userSection, ns, "msvSDContainerMargin"));
        }

        private static void AssertContainerMembershipXml(string filePath, string containerText, string[] expectedMemberTexts, string[] absentMemberTexts) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument page = ReadXml(archive, "visio/pages/page1.xml");

            XElement container = ShapeByText(page, ns, containerText);
            string containerId = container.Attribute("ID")!.Value;
            string containerFormula = RelationshipFormula(container, ns) ?? string.Empty;

            foreach (string memberText in expectedMemberTexts) {
                XElement member = ShapeByText(page, ns, memberText);
                string memberId = member.Attribute("ID")!.Value;
                Assert.Contains($"DEPENDSON(1,Sheet.{memberId}!SheetRef())", containerFormula);
                Assert.Contains($"DEPENDSON(4,Sheet.{containerId}!SheetRef())", RelationshipFormula(member, ns) ?? string.Empty);
            }

            foreach (string memberText in absentMemberTexts) {
                XElement member = ShapeByText(page, ns, memberText);
                string memberId = member.Attribute("ID")!.Value;
                Assert.DoesNotContain($"DEPENDSON(1,Sheet.{memberId}!SheetRef())", containerFormula);
                Assert.DoesNotContain($"DEPENDSON(4,Sheet.{containerId}!SheetRef())", RelationshipFormula(member, ns) ?? string.Empty);
            }
        }

        private static void AssertContainerMetadataXml(
            string filePath,
            string containerText,
            string expectedMargin,
            string expectedHeadingHeight,
            string expectedResize,
            string expectedLocked,
            string expectedNoHighlight,
            string expectedNoRibbon,
            string expectedContainerStyle,
            string expectedHeadingStyle) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument page = ReadXml(archive, "visio/pages/page1.xml");

            XElement container = ShapeByText(page, ns, containerText);
            XElement userSection = container.Elements(ns + "Section")
                .Single(section => (string?)section.Attribute("N") == "User");

            Assert.Equal(expectedMargin, UserCellValue(userSection, ns, "msvSDContainerMargin"));
            Assert.Equal(expectedHeadingHeight, UserCellValue(userSection, ns, VisioSemanticUserCells.ContainerHeadingHeight));
            Assert.Equal(expectedResize, UserCellValue(userSection, ns, "msvSDContainerResize"));
            Assert.Equal(expectedLocked, UserCellValue(userSection, ns, "msvSDContainerLocked"));
            Assert.Equal(expectedNoHighlight, UserCellValue(userSection, ns, "msvSDContainerNoHighlight"));
            Assert.Equal(expectedNoRibbon, UserCellValue(userSection, ns, "msvSDContainerNoRibbon"));
            Assert.Equal(expectedContainerStyle, UserCellValue(userSection, ns, "msvSDContainerStyle"));
            Assert.Equal(expectedHeadingStyle, UserCellValue(userSection, ns, "msvSDHeadingStyle"));
        }

        private static XElement ShapeByText(XDocument page, XNamespace ns, string text) {
            return page.Descendants(ns + "Shape")
                .Single(shape => shape.Element(ns + "Text")?.Value == text);
        }

        private static string? RelationshipFormula(XElement shape, XNamespace ns) {
            return shape.Elements(ns + "Cell")
                .SingleOrDefault(cell => (string?)cell.Attribute("N") == "Relationships")
                ?.Attribute("F")
                ?.Value;
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
