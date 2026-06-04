using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests {
    public class VisioShapeDataTests {
        [Fact]
        public void TypedShapeDataSavesLoadsPreservesMetadataAndKeepsDictionaryCompatibility() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Shape Data", 8.5, 6);
            VisioShape server = page.AddRectangle(2.5, 4, 2.2, 1, "Server");
            server.SetShapeData("Owner", "Operations", "Owner", VisioShapeDataType.String, "Owning support team");
            VisioShapeDataRow cost = server.SetShapeData("MonthlyCost", "1250", "Monthly cost", VisioShapeDataType.Currency, "Estimated monthly cost", "$#,##0");
            cost.SortKey = "020";
            cost.Verify = true;

            page.AddRectangle(6, 4, 2.2, 1, "Database")
                .SetShapeData("Owner", "Data", "Owner", VisioShapeDataType.String, "Owning support team");

            page.SelectWithShapeData("Owner", "Operations")
                .Fill(Color.LightBlue)
                .ShapeData("Reviewed", "Yes", "Reviewed", VisioShapeDataType.Boolean, "Architecture review complete");

            Assert.Single(page.ShapesWithShapeData("Owner", "Operations"));
            Assert.Single(page.ShapesWithData("Reviewed", "Yes"));

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertShapeDataXml(filePath, "Operations", "Yes");

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape loadedServer = loaded.Pages[0].Shapes.Single(shape => shape.Text == "Server");
            Assert.Equal("Operations", loadedServer.GetShapeDataValue("Owner"));
            Assert.Equal("Owner", loadedServer.FindShapeData("Owner")!.Label);
            Assert.Equal("Owning support team", loadedServer.FindShapeData("Owner")!.Prompt);
            Assert.Equal(VisioShapeDataType.Currency, loadedServer.FindShapeData("MonthlyCost")!.Type);
            Assert.Equal("$#,##0", loadedServer.FindShapeData("MonthlyCost")!.Format);
            Assert.True(loadedServer.FindShapeData("MonthlyCost")!.Verify);

            loadedServer.Data["Owner"] = "Platform";
            Assert.Equal("Platform", loadedServer.GetShapeDataValue("Owner"));
            loaded.Save(updatedPath);

            Assert.Empty(VisioValidator.Validate(updatedPath));
            AssertShapeDataXml(updatedPath, "Platform", "Yes");
        }

        [Fact]
        public void ConnectorShapeDataSetReusesExistingRowNameCasing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Connector Data", 8.5, 6);
            VisioShape source = page.AddRectangle(2, 4, 1.5, 0.75, "Source");
            VisioShape target = page.AddRectangle(5, 4, 1.5, 0.75, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic);
            connector.Label = "route";
            connector.SetShapeData("Owner", "Operations", "Owner", VisioShapeDataType.String);
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector loadedConnector = loaded.Pages[0].Connectors.Single();
            loadedConnector.SetShapeData("owner", "Platform");
            Assert.Equal("Platform", loadedConnector.Data["Owner"]);
            Assert.False(loadedConnector.Data.ContainsKey("owner"));
            loaded.Save(updatedPath);

            Assert.Empty(VisioValidator.Validate(updatedPath));
            AssertConnectorShapeDataXml(updatedPath, "Owner", "Platform");
        }

        [Fact]
        public void ConnectorShapeDataClearOverridesPreservedLoadedValue() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Connector Data", 8.5, 6);
            VisioShape source = page.AddRectangle(2, 4, 1.5, 0.75, "Source");
            VisioShape target = page.AddRectangle(5, 4, 1.5, 0.75, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic);
            connector.SetShapeData("Owner", "Operations", "Owner", VisioShapeDataType.String);
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector loadedConnector = loaded.Pages[0].Connectors.Single();
            loadedConnector.SetShapeData("Owner", null);
            Assert.Equal(string.Empty, loadedConnector.GetShapeDataValue("Owner"));
            Assert.False(loadedConnector.Data.ContainsKey("Owner"));
            loaded.Save(updatedPath);

            Assert.Empty(VisioValidator.Validate(updatedPath));
            AssertConnectorShapeDataXml(updatedPath, "Owner", string.Empty);
        }

        [Fact]
        public void ConnectorShapeDataRowValueEditWinsOverStaleDictionaryMirror() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Connector Data", 8.5, 6);
            VisioShape source = page.AddRectangle(2, 4, 1.5, 0.75, "Source");
            VisioShape target = page.AddRectangle(5, 4, 1.5, 0.75, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic);
            VisioShapeDataRow owner = connector.SetShapeData("Owner", "Operations", "Owner", VisioShapeDataType.String);
            owner.Value = "Platform";

            Assert.Equal("Platform", connector.GetShapeDataValue("Owner"));
            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertConnectorShapeDataXml(filePath, "Owner", "Platform");
        }

        [Fact]
        public void ShapeDataSchemaAppliesValidatesAndRoundTripsRows() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioShapeDataSchema schema = VisioShapeDataSchema.Create()
                .Field("Owner", "Owner", VisioShapeDataType.String, "Unassigned", "Owning team", sortKey: "010", required: true)
                .Field("Risk", "Risk", VisioShapeDataType.FixedList, "Medium", "Operational risk", sortKey: "020", required: true, verify: true, allowedValues: new[] { "Low", "Medium", "High" })
                .Field("MonthlyCost", "Monthly cost", VisioShapeDataType.Currency, "0", "Estimated monthly run cost", "$#,##0", "030", invisible: false);

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Schema", 8.5, 6);
            VisioShape server = page.AddRectangle(2, 4, 1.5, 0.75, "Server");
            VisioShape api = page.AddRectangle(5, 4, 1.5, 0.75, "API");
            server.SetShapeData("Owner", "Operations");

            schema.ApplyTo(server);
            page.SelectShapes(shape => shape.Text == "API").ShapeData(schema);
            VisioConnector connector = page.AddConnector(server, api, ConnectorKind.Dynamic);
            page.SelectConnectors(current => current == connector).ShapeData(schema, overwriteValues: true);

            Assert.Equal("Operations", server.GetShapeDataValue("Owner"));
            Assert.Equal("Unassigned", api.GetShapeDataValue("Owner"));
            Assert.Equal("Medium", api.GetShapeDataValue("Risk"));
            Assert.Equal("Medium", connector.GetShapeDataValue("Risk"));
            Assert.Empty(schema.Validate(server));
            Assert.Empty(schema.Validate(connector));

            api.SetShapeData("Risk", "Critical");
            VisioShapeDataSchemaIssue issue = Assert.Single(schema.Validate(api));
            Assert.Equal(VisioShapeDataSchemaIssueKind.ValueNotAllowed, issue.Kind);
            Assert.Equal("Risk", issue.FieldName);

            api.SetShapeData("Risk", "High");
            Assert.Empty(schema.Validate(api));

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertShapeDataSchemaXml(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape loadedServer = loaded.Pages[0].Shapes.Single(shape => shape.Text == "Server");
            VisioShape loadedApi = loaded.Pages[0].Shapes.Single(shape => shape.Text == "API");
            VisioConnector loadedConnector = loaded.Pages[0].Connectors.Single();
            Assert.Equal("Operations", loadedServer.GetShapeDataValue("Owner"));
            Assert.Equal("High", loadedApi.GetShapeDataValue("Risk"));
            Assert.Equal("Medium", loadedConnector.GetShapeDataValue("Risk"));
            Assert.Equal("Low;Medium;High", loadedServer.FindShapeData("Risk")!.Format);
            Assert.Equal("020", loadedServer.FindShapeData("Risk")!.SortKey);
            Assert.True(loadedServer.FindShapeData("Risk")!.Verify);
            Assert.False(loadedServer.FindShapeData("MonthlyCost")!.Invisible);
        }

        [Fact]
        public void ShapeDataSchemaPreservesLegacyDataValuesByDefault() {
            VisioShapeDataSchema schema = VisioShapeDataSchema.Create()
                .Field("Owner", "Owner", VisioShapeDataType.String, "Unassigned");
            VisioShape shape = new("shape", 0, 0, 1, 1, "Shape");
            shape.Data["Owner"] = "Ops";

            schema.ApplyTo(shape);

            Assert.Equal("Ops", shape.GetShapeDataValue("Owner"));
            Assert.Equal("Ops", shape.FindShapeData("Owner")!.Value);
        }

        [Fact]
        public void DataGraphicsCreateVisibleShapeDataAdornments() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            document.UseMastersByDefault = true;
            VisioPage page = document.AddPage("Data Graphics", 8.5, 6);
            VisioShape api = page.AddRectangle(2, 4, 1.5, 0.75, "API");
            VisioShape database = page.AddRectangle(5, 4, 1.5, 0.75, "Database");
            api.SetShapeData("Status", "Healthy", "Status", VisioShapeDataType.FixedList, format: "Healthy;Warning;Critical");
            api.SetShapeData("Slo", "72", "SLO", VisioShapeDataType.Number);
            database.SetShapeData("Status", "Warning", "Status", VisioShapeDataType.FixedList, format: "Healthy;Warning;Critical");
            database.SetShapeData("Slo", "41", "SLO", VisioShapeDataType.Number);

            VisioDataGraphic graphic = VisioDataGraphic.Create()
                .Badge("Status")
                .Bar("Slo", maximumValue: 100, label: "SLO");

            IReadOnlyList<VisioShape> generated = page.SelectWithShapeData("Status", value => !string.IsNullOrWhiteSpace(value))
                .AddDataGraphics(graphic);

            Assert.Equal(8, generated.Count);
            Assert.All(generated, shape => Assert.True(shape.IsDiagramAdornment));
            Assert.All(generated, shape => Assert.Equal("Data Graphics", shape.LayerNames.Single()));
            Assert.Contains(generated, shape => shape.Text == "Status: Healthy");
            Assert.Contains(generated, shape => shape.Text == "SLO: 72");

            VisioShape apiBarFill = generated.Single(shape =>
                shape.GetUserCellValue("OfficeIMO.DataGraphicTargetId") == api.Id &&
                shape.GetUserCellValue("OfficeIMO.DataGraphicField") == "Slo" &&
                shape.GetUserCellValue("OfficeIMO.DataGraphicRole") == "BarFill");
            Assert.Equal(0.792D, apiBarFill.Width, 3);
            Assert.Equal("0.72", apiBarFill.GetShapeDataValue("Percent"));

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            using (ZipArchive archive = ZipFile.OpenRead(filePath)) {
                XDocument pageXml = ReadXml(archive, "visio/pages/page1.xml");
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement[] graphicShapeElements = pageXml.Descendants(ns + "Shape")
                    .Where(shape => ((string?)shape.Attribute("NameU"))?.StartsWith("OfficeIMO Data Graphic", StringComparison.Ordinal) == true)
                    .ToArray();

                Assert.Equal(6, graphicShapeElements.Length);
                Assert.All(graphicShapeElements, shape => Assert.Null(shape.Attribute("Master")));
                Assert.All(graphicShapeElements, shape => Assert.Contains(shape.Elements(ns + "Section"), section => (string?)section.Attribute("N") == "Geometry"));
            }

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape loadedFill = loaded.Pages[0].Shapes.Single(shape =>
                shape.GetUserCellValue("OfficeIMO.DataGraphicTargetId") == api.Id &&
                shape.GetUserCellValue("OfficeIMO.DataGraphicField") == "Slo" &&
                shape.GetUserCellValue("OfficeIMO.DataGraphicRole") == "BarFill");
            Assert.True(loadedFill.IsDiagramAdornment);
            Assert.Equal("72", loadedFill.GetShapeDataValue("DataGraphicValue"));
            Assert.Equal("0.72", loadedFill.GetShapeDataValue("Percent"));
        }

        [Fact]
        public void DataGraphicsUsePageUnitsForGeneratedAdornmentGeometry() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Metric Data Graphics", 12, 8, VisioMeasurementUnit.Centimeters);
            VisioShape service = page.AddRectangle(4, 5, 2, 1, "Service");
            service.SetShapeData("Status", "Watch", "Status", VisioShapeDataType.String);
            service.SetShapeData("Slo", "80", "SLO", VisioShapeDataType.Number);

            VisioDataGraphic graphic = VisioDataGraphic.Create();
            graphic.Gap = 0.5D;
            graphic.ItemSpacing = 0.8D;
            graphic.Badge("Status");
            graphic.Bar("Slo", maximumValue: 100, label: "SLO");

            VisioShape[] generated = page.AddDataGraphics(service, graphic).ToArray();
            VisioShape badge = generated.Single(shape => shape.GetUserCellValue("OfficeIMO.DataGraphicRole") == "Badge");
            VisioShape bar = generated.Single(shape => shape.GetUserCellValue("OfficeIMO.DataGraphicRole") == "BarBackground");

            Assert.Equal(1.15D / 2.54D, badge.Width, 3);
            Assert.Equal(0.28D / 2.54D, badge.Height, 3);
            Assert.Equal(1.1D / 2.54D, bar.Width, 3);
            Assert.Equal(0.12D / 2.54D, bar.Height, 3);
            Assert.Equal(0.5D / 2.54D, badge.PinX - service.PinX - (service.Width / 2D) - (badge.Width / 2D), 3);
            Assert.Equal(0.8D / 2.54D, badge.PinY - bar.PinY, 3);
        }

        [Fact]
        public void DataGraphicsUsePageCoordinatesForGroupedChildren() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Grouped Data Graphics", 8.5, 6);
            VisioShape group = new("group") {
                Type = "Group",
                PinX = 5,
                PinY = 4,
                Width = 4,
                Height = 3,
                LocPinX = 2,
                LocPinY = 1.5
            };
            VisioShape child = new("child", 1.25, 1.25, 1, 0.6, "Child");
            child.SetShapeData("Status", "Healthy", "Status", VisioShapeDataType.String);
            group.Children.Add(child);
            page.Shapes.Add(group);

            VisioDataGraphic graphic = VisioDataGraphic.Create();
            graphic.Gap = 0.2D;
            VisioShape badge = Assert.Single(page.AddDataGraphics(child, graphic.Badge("Status")));

            double childPageX = group.PinX - group.LocPinX + child.PinX;
            double childPageY = group.PinY - group.LocPinY + child.PinY;
            Assert.Equal(childPageX + (child.Width / 2D) + 0.2D + (badge.Width / 2D), badge.PinX, 3);
            Assert.Equal(childPageY, badge.PinY, 3);
        }

        [Fact]
        public void DataGraphicsRejectInvalidDataBarRangesBeforeGeneratingGeometry() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Invalid Data Graphic", 8.5, 6);
            VisioShape target = page.AddRectangle(2, 4, 1.5, 0.75, "Service");
            target.SetShapeData("Slo", "80", "SLO", VisioShapeDataType.Number);

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                page.AddDataGraphics(target, VisioDataGraphic.Create().Bar("Slo", minimumValue: 100, maximumValue: 100)));

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                page.AddDataGraphics(target, VisioDataGraphic.Create().Bar("Slo", minimumValue: double.NaN, maximumValue: 100)));

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                page.AddDataGraphics(target, VisioDataGraphic.Create().Bar("Slo", minimumValue: 0, maximumValue: double.PositiveInfinity)));

            Assert.Single(page.Shapes);
        }

        [Fact]
        public void PremiumGalleryPromotesDataGraphicsIntoExecutiveScenario() {
            string folderPath = Path.Combine(Path.GetTempPath(), "OfficeIMO-Visio-Premium-DataGraphics-" + Guid.NewGuid().ToString("N"));
            try {
                IReadOnlyList<VisioGalleryResult> results = VisioPremiumGallery.Create(folderPath);
                VisioGalleryResult executive = results.Single(result => result.Name == "Premium Executive Dependencies");

                Assert.True(executive.IsClean, string.Join(Environment.NewLine, executive.PackageIssues.Concat(executive.QualityIssues.Select(issue => issue.ToString()))));

                VisioDocument loaded = VisioDocument.Load(executive.FilePath);
                VisioPage page = loaded.Pages.Single();
                VisioShape target = page.FindShapeById("warehouse")!;
                Assert.NotNull(target);
                Assert.Equal("Watch", target.GetShapeDataValue("Status"));
                Assert.Equal("91", target.GetShapeDataValue("Slo"));

                VisioShape[] dataGraphics = page.Shapes
                    .Where(shape => shape.GetUserCellValue("OfficeIMO.DataGraphicTargetId") == target.Id)
                    .OrderBy(shape => shape.GetUserCellValue("OfficeIMO.DataGraphicRole"), StringComparer.Ordinal)
                    .ToArray();

                Assert.Equal(4, dataGraphics.Length);
                Assert.All(dataGraphics, shape => Assert.True(shape.IsDiagramAdornment));
                Assert.Contains(dataGraphics, shape => shape.Text == "Status: Watch");
                Assert.Contains(dataGraphics, shape => shape.Text == "SLO: 91");
                Assert.Contains(dataGraphics, shape => shape.GetUserCellValue("OfficeIMO.DataGraphicRole") == "BarFill" && shape.GetShapeDataValue("Percent") == "0.91");
            } finally {
                if (Directory.Exists(folderPath)) {
                    Directory.Delete(folderPath, recursive: true);
                }
            }
        }

        [Fact]
        public void ConnectorShapeDataSetClearsLoadedValueFormula() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Connector Data", 8.5, 6);
            VisioShape source = page.AddRectangle(2, 4, 1.5, 0.75, "Source");
            VisioShape target = page.AddRectangle(5, 4, 1.5, 0.75, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic);
            connector.SetShapeData("Owner", "Operations", "Owner", VisioShapeDataType.String);
            document.Save();
            SetConnectorValueCellToFormulaOnly(filePath, "Owner", "\"Operations\"");

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector loadedConnector = loaded.Pages[0].Connectors.Single();
            loadedConnector.SetShapeData("Owner", "Platform");
            loaded.Save(updatedPath);

            Assert.Empty(VisioValidator.Validate(updatedPath));
            AssertConnectorShapeDataValueHasNoFormula(updatedPath, "Owner", "Platform");
        }

        private static void AssertShapeDataXml(string filePath, string ownerValue, string reviewedValue) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument page = ReadXml(archive, "visio/pages/page1.xml");
            XElement server = page.Descendants(ns + "Shape")
                .Single(shape => shape.Element(ns + "Text")?.Value == "Server");
            XElement propSection = server.Elements(ns + "Section")
                .Single(section => (string?)section.Attribute("N") == "Prop");

            XElement owner = Row(propSection, ns, "Owner");
            Assert.Equal(ownerValue, CellValue(owner, ns, "Value"));
            Assert.Equal("Owner", CellValue(owner, ns, "Label"));
            Assert.Equal("Owning support team", CellValue(owner, ns, "Prompt"));
            Assert.Equal(((int)VisioShapeDataType.String).ToString(), CellValue(owner, ns, "Type"));

            XElement cost = Row(propSection, ns, "MonthlyCost");
            Assert.Equal("1250", CellValue(cost, ns, "Value"));
            Assert.Equal(((int)VisioShapeDataType.Currency).ToString(), CellValue(cost, ns, "Type"));
            Assert.Equal("$#,##0", CellValue(cost, ns, "Format"));
            Assert.Equal("1", CellValue(cost, ns, "Verify"));

            XElement reviewed = Row(propSection, ns, "Reviewed");
            Assert.Equal(reviewedValue, CellValue(reviewed, ns, "Value"));
            Assert.Equal("Reviewed", CellValue(reviewed, ns, "Label"));
            Assert.Equal(((int)VisioShapeDataType.Boolean).ToString(), CellValue(reviewed, ns, "Type"));
        }

        private static void AssertShapeDataSchemaXml(string filePath) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument page = ReadXml(archive, "visio/pages/page1.xml");
            XElement server = page.Descendants(ns + "Shape")
                .Single(shape => shape.Element(ns + "Text")?.Value == "Server");
            XElement propSection = server.Elements(ns + "Section")
                .Single(section => (string?)section.Attribute("N") == "Prop");

            XElement owner = Row(propSection, ns, "Owner");
            Assert.Equal("Operations", CellValue(owner, ns, "Value"));
            Assert.Equal("Owner", CellValue(owner, ns, "Label"));
            Assert.Equal("Owning team", CellValue(owner, ns, "Prompt"));
            Assert.Equal("010", CellValue(owner, ns, "SortKey"));

            XElement risk = Row(propSection, ns, "Risk");
            Assert.Equal("Medium", CellValue(risk, ns, "Value"));
            Assert.Equal(((int)VisioShapeDataType.FixedList).ToString(), CellValue(risk, ns, "Type"));
            Assert.Equal("Low;Medium;High", CellValue(risk, ns, "Format"));
            Assert.Equal("020", CellValue(risk, ns, "SortKey"));
            Assert.Equal("1", CellValue(risk, ns, "Verify"));

            XElement cost = Row(propSection, ns, "MonthlyCost");
            Assert.Equal("0", CellValue(cost, ns, "Value"));
            Assert.Equal(((int)VisioShapeDataType.Currency).ToString(), CellValue(cost, ns, "Type"));
            Assert.Equal("$#,##0", CellValue(cost, ns, "Format"));
            Assert.Equal("030", CellValue(cost, ns, "SortKey"));
            Assert.Equal("0", CellValue(cost, ns, "Invisible"));
        }

        private static void AssertConnectorShapeDataXml(string filePath, string rowName, string ownerValue) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument page = ReadXml(archive, "visio/pages/page1.xml");
            XElement connector = page.Descendants(ns + "Shape")
                .Single(shape => shape.Elements(ns + "Section")
                    .Any(section => (string?)section.Attribute("N") == "Prop" &&
                        section.Elements(ns + "Row")
                            .Any(row => string.Equals((string?)row.Attribute("N"), rowName, StringComparison.Ordinal))));
            XElement propSection = connector.Elements(ns + "Section")
                .Single(section => (string?)section.Attribute("N") == "Prop");

            Assert.Single(propSection.Elements(ns + "Row"),
                row => string.Equals((string?)row.Attribute("N"), rowName, StringComparison.Ordinal));
            Assert.DoesNotContain(propSection.Elements(ns + "Row"),
                row => string.Equals((string?)row.Attribute("N"), rowName.ToLowerInvariant(), StringComparison.Ordinal));
            Assert.Equal(ownerValue, CellValue(Row(propSection, ns, rowName), ns, "Value"));
        }

        private static void AssertConnectorShapeDataValueHasNoFormula(string filePath, string rowName, string ownerValue) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument page = ReadXml(archive, "visio/pages/page1.xml");
            XElement propSection = page.Descendants(ns + "Section")
                .Single(section => (string?)section.Attribute("N") == "Prop" &&
                    section.Elements(ns + "Row").Any(row => (string?)row.Attribute("N") == rowName));
            XElement valueCell = Row(propSection, ns, rowName).Elements(ns + "Cell")
                .Single(cell => (string?)cell.Attribute("N") == "Value");

            Assert.Equal(ownerValue, valueCell.Attribute("V")?.Value);
            Assert.Null(valueCell.Attribute("F"));
        }

        private static XElement Row(XElement section, XNamespace ns, string name) {
            return section.Elements(ns + "Row")
                .Single(row => (string?)row.Attribute("N") == name);
        }

        private static string CellValue(XElement row, XNamespace ns, string cellName) {
            return row.Elements(ns + "Cell")
                .Single(cell => (string?)cell.Attribute("N") == cellName)
                .Attribute("V")!
                .Value;
        }

        private static void SetConnectorValueCellToFormulaOnly(string filePath, string rowName, string formula) {
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            using ZipArchive archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            ZipArchiveEntry entry = archive.GetEntry("visio/pages/page1.xml") ?? throw new InvalidOperationException("Missing visio/pages/page1.xml");
            XDocument page;
            using (Stream stream = entry.Open()) {
                page = XDocument.Load(stream);
            }

            XElement propSection = page.Descendants(ns + "Section")
                .Single(section => (string?)section.Attribute("N") == "Prop" &&
                    section.Elements(ns + "Row").Any(row => (string?)row.Attribute("N") == rowName));
            XElement valueCell = Row(propSection, ns, rowName).Elements(ns + "Cell")
                .Single(cell => (string?)cell.Attribute("N") == "Value");
            valueCell.Attribute("V")?.Remove();
            valueCell.SetAttributeValue("F", formula);

            entry.Delete();
            ZipArchiveEntry updated = archive.CreateEntry("visio/pages/page1.xml");
            using Stream output = updated.Open();
            page.Save(output);
        }

        private static XDocument ReadXml(ZipArchive archive, string entryName) {
            ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Missing " + entryName);
            using Stream stream = entry.Open();
            return XDocument.Load(stream);
        }
    }
}
