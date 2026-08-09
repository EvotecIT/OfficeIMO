using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_InspectionCountsSlicerAndTimelinePackageParts() {
            string filePath = Path.Combine(_directoryWithFiles, "PackageInteractions.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Value");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                WriteExtendedPart(
                    workbookPart.AddExtendedPart(
                        "http://schemas.microsoft.com/office/2007/relationships/slicerCache",
                        "application/vnd.ms-excel.slicerCache+xml",
                        ".xml"),
                    "<slicerCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"/>");
                WriteExtendedPart(
                    workbookPart.AddExtendedPart(
                        "http://schemas.microsoft.com/office/2011/relationships/timelineCache",
                        "application/vnd.ms-excel.timelineCache+xml",
                        ".xml"),
                    "<timelineCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2011/1/timeline\"/>");
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
                Assert.Equal(1, snapshot.SlicerPartCount);
                Assert.Equal(1, snapshot.TimelinePartCount);
                Assert.True(snapshot.HasSlicers);
                Assert.True(snapshot.HasTimelines);
                Assert.False(snapshot.HasSlicerBindingMetadata);
                Assert.False(snapshot.HasTimelineBindingMetadata);
                Assert.Empty(document.GetWorkbookSlicerCaches());
                Assert.Empty(document.GetWorkbookTimelineCaches());
            }
        }

        [Fact]
        public void Test_LegacyOfficeImoPivotInteractionMetadata_IsRecognizedWithoutMaskingNativeCaches() {
            string filePath = Path.Combine(_directoryWithFiles, "PackageInteractions.LegacyOfficeImoMetadata.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Value");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                WriteExtendedPart(
                    workbookPart.AddExtendedPart(
                        "http://schemas.microsoft.com/office/2007/relationships/slicerCache",
                        "application/vnd.ms-excel.slicerCache+xml",
                        "xml"),
                    "<slicerCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" name=\"LegacyRegion\" sourceName=\"Region\" pivotTableName=\"SalesPivot\"/>");
                WriteExtendedPart(
                    workbookPart.AddExtendedPart(
                        "http://schemas.microsoft.com/office/2007/relationships/slicerCache",
                        "application/vnd.ms-excel.slicerCache+xml",
                        "xml"),
                    "<slicerCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" name=\"LegacyRegionNoPivot\" sourceName=\"RegionNoPivot\"/>");
                WriteExtendedPart(
                    workbookPart.AddExtendedPart(
                        "http://schemas.microsoft.com/office/2007/relationships/slicerCache",
                        "application/vnd.ms-excel.slicerCache+xml",
                        "xml"),
                    "<slicerCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" name=\"NativeRegion\" sourceName=\"NativeRegionSource\"><data/></slicerCacheDefinition>");
                WriteExtendedPart(
                    workbookPart.AddExtendedPart(
                        "http://schemas.microsoft.com/office/2011/relationships/timelineCache",
                        "application/vnd.ms-excel.timelineCache+xml",
                        "xml"),
                    "<timelineCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2011/1/main\" name=\"LegacyOrderDate\" sourceName=\"OrderDate\" pivotTableName=\"SalesPivot\"/>");
                WriteExtendedPart(
                    workbookPart.AddExtendedPart(
                        "http://schemas.microsoft.com/office/2011/relationships/timelineCache",
                        "application/vnd.ms-excel.timelineCache+xml",
                        "xml"),
                    "<timelineCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2011/1/main\" name=\"LegacyOrderDateNoPivot\" sourceName=\"OrderDateNoPivot\"/>");
                WriteExtendedPart(
                    workbookPart.AddExtendedPart(
                        "http://schemas.microsoft.com/office/2011/relationships/timelineCache",
                        "application/vnd.ms-excel.timelineCache+xml",
                        "xml"),
                    "<timelineCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2011/1/main\" name=\"NativeOrderDate\" sourceName=\"NativeOrderDateSource\"><state/></timelineCacheDefinition>");
            }

            using (var document = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
                IReadOnlyList<ExcelPivotInteractionCacheInfo> slicers = document.GetWorkbookSlicerCaches();
                Assert.Equal(2, slicers.Count);
                ExcelPivotInteractionCacheInfo slicer = Assert.Single(slicers, cache => cache.Name == "LegacyRegion");
                Assert.Equal("LegacyRegion", slicer.Name);
                Assert.Equal("Region", slicer.SourceName);
                Assert.Equal("SalesPivot", slicer.PivotTableName);
                ExcelPivotInteractionCacheInfo slicerWithoutPivot = Assert.Single(slicers, cache => cache.Name == "LegacyRegionNoPivot");
                Assert.Equal("RegionNoPivot", slicerWithoutPivot.SourceName);
                Assert.Null(slicerWithoutPivot.PivotTableName);

                IReadOnlyList<ExcelPivotInteractionCacheInfo> timelines = document.GetWorkbookTimelineCaches();
                Assert.Equal(2, timelines.Count);
                ExcelPivotInteractionCacheInfo timeline = Assert.Single(timelines, cache => cache.Name == "LegacyOrderDate");
                Assert.Equal("LegacyOrderDate", timeline.Name);
                Assert.Equal("OrderDate", timeline.SourceName);
                Assert.Equal("SalesPivot", timeline.PivotTableName);
                ExcelPivotInteractionCacheInfo timelineWithoutPivot = Assert.Single(timelines, cache => cache.Name == "LegacyOrderDateNoPivot");
                Assert.Equal("OrderDateNoPivot", timelineWithoutPivot.SourceName);
                Assert.Null(timelineWithoutPivot.PivotTableName);

                ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
                Assert.Equal(1, snapshot.SlicerPartCount);
                Assert.Equal(1, snapshot.TimelinePartCount);
                Assert.Equal(2, snapshot.SlicerBindingMetadataPartCount);
                Assert.Equal(2, snapshot.TimelineBindingMetadataPartCount);
            }
        }

        [Fact]
        public void Test_CopyPackage_PreservesPartsAndNormalizesWorkbookContentType() {
            string sourcePath = Path.Combine(_directoryWithFiles, "PackageClone.Source.xlsx");
            string destinationPath = Path.Combine(_directoryWithFiles, "PackageClone.Target.xlsm");

            using (var document = ExcelDocument.Create(sourcePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Value");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(sourcePath, true)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                WriteExtendedPart(
                    workbookPart.AddExtendedPart(
                        "http://schemas.microsoft.com/office/2007/relationships/slicerCache",
                        "application/vnd.ms-excel.slicerCache+xml",
                        "xml"),
                    "<slicerCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"/>");
                WriteExtendedPart(
                    workbookPart.AddExtendedPart(
                        "http://schemas.microsoft.com/office/2011/relationships/timelineCache",
                        "application/vnd.ms-excel.timelineCache+xml",
                        "xml"),
                    "<timelineCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2011/1/timeline\"/>");
            }

            ExcelDocument.CopyPackage(sourcePath, destinationPath);

            Assert.Equal(
                "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                GetWorkbookOverrideContentType(destinationPath));

            using (var document = ExcelDocument.Load(destinationPath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                var worksheetPart = document._spreadSheetDocument.WorkbookPart!.WorksheetParts.Single();
                Assert.Equal("Value", GetCellValue(document._spreadSheetDocument, worksheetPart, "A1"));

                ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
                Assert.Equal(1, snapshot.SlicerPartCount);
                Assert.Equal(1, snapshot.TimelinePartCount);
            }
        }

        [Fact]
        public void Test_InspectionSnapshot_SkipsNonWorksheetSheetParts() {
            string filePath = Path.Combine(_directoryWithFiles, "PackageInteractions.ChartSheetInspection.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Value");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                ChartsheetPart chartsheetPart = workbookPart.AddNewPart<ChartsheetPart>();
                chartsheetPart.Chartsheet = new Chartsheet(new SheetViews(new SheetView { WorkbookViewId = 0U }));
                chartsheetPart.Chartsheet.Save();

                Sheets sheets = workbookPart.Workbook.Sheets!;
                uint nextSheetId = sheets.Elements<Sheet>().Select(sheet => sheet.SheetId?.Value ?? 0U).Max() + 1U;
                sheets.Append(new Sheet {
                    Id = workbookPart.GetIdOfPart(chartsheetPart),
                    SheetId = nextSheetId,
                    Name = "Chart View"
                });
                workbookPart.Workbook.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
                ExcelWorksheetSnapshot worksheet = Assert.Single(snapshot.Worksheets);
                Assert.Equal("Data", worksheet.Name);
            }
        }


        [Fact]
        public void Test_CopyPackage_RejectsMacroEnabledSourceToMacroFreeDestination() {
            string sourcePath = Path.Combine(_directoryWithFiles, "PackageClone.MacroSource.xlsx");
            string macroPath = Path.Combine(_directoryWithFiles, "PackageClone.MacroSource.xlsm");
            string destinationPath = Path.Combine(_directoryWithFiles, "PackageClone.MacroBlocked.xlsx");

            using (var document = ExcelDocument.Create(sourcePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Value");
                document.Save();
            }

            ExcelDocument.CopyPackage(sourcePath, macroPath);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                ExcelDocument.CopyPackage(macroPath, destinationPath));
            Assert.Contains("Macro-enabled workbook packages", exception.Message);
            Assert.False(File.Exists(destinationPath));
        }

        [Fact]
        public void Test_ConnectionAndQueryTableMetadataParts_AreAuthoredAndInspected() {
            string filePath = Path.Combine(_directoryWithFiles, "PackageInteractions.ConnectionMetadata.xlsx");
            const string connectionXml = "<connections xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\"><connection id=\"1\" name=\"SalesConnection\" type=\"5\" refreshedVersion=\"7\"/></connections>";
            const string queryTableXml = "<queryTable xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"SalesQuery\" connectionId=\"1\"/>";

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Region");
                document.AddWorkbookConnectionMetadata(connectionXml);
                document.AddWorksheetQueryTableMetadata("Data", queryTableXml);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
                Assert.Equal(1, snapshot.ConnectionPartCount);
                Assert.Equal(1, snapshot.QueryTablePartCount);
                Assert.True(snapshot.HasConnections);
                Assert.True(snapshot.HasQueryTables);
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                Assert.Contains("SalesConnection", ReadSinglePackagePartText(workbookPart, "connections"));

                var worksheetPart = workbookPart.WorksheetParts.Single();
                Assert.Contains("SalesQuery", ReadSinglePackagePartText(worksheetPart, "queryTable"));
            }
        }

        [Fact]
        public void Test_ConnectionMetadata_MergesIntoTypedWorkbookConnectionsPart() {
            string filePath = Path.Combine(_directoryWithFiles, "PackageInteractions.TypedConnectionMetadata.xlsx");
            const string connectionXml = "<connections xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\"><connection id=\"2\" name=\"Added\" type=\"5\" refreshedVersion=\"7\"/></connections>";

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Data");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                ConnectionsPart connectionsPart = spreadsheet.WorkbookPart!.AddNewPart<ConnectionsPart>();
                connectionsPart.Connections = new Connections(
                    new Connection { Id = 1U, Name = "Existing", Type = 5, RefreshedVersion = 7 });
                connectionsPart.Connections.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                ExcelPackagePartInfo part = document.AddWorkbookConnectionMetadata(connectionXml);
                Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml", part.ContentType);
                Assert.IsType<ConnectionsPart>(document.WorkbookPartRoot.GetPartById(part.RelationshipId));
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                ConnectionsPart connectionsPart = Assert.Single(workbookPart.GetPartsOfType<ConnectionsPart>());
                Assert.Contains(connectionsPart.Connections!.Elements<Connection>(), connection => connection.Name?.Value == "Existing");
                Assert.Contains(connectionsPart.Connections!.Elements<Connection>(), connection => connection.Name?.Value == "Added");
                Assert.DoesNotContain(workbookPart.Parts.Select(pair => pair.OpenXmlPart), part => part is ExtendedPart && part.ContentType.IndexOf("connections", StringComparison.OrdinalIgnoreCase) >= 0);
            }
        }

        [Fact]
        public void Test_SlicerAndTimelineMetadataParts_AreAuthoredAndInspected() {
            string filePath = Path.Combine(_directoryWithFiles, "PackageInteractions.SlicerTimelineMetadata.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Region");
                document.AddWorkbookSlicerCache(new ExcelSlicerCacheOptions {
                    Name = "RegionSlicer",
                    SourceName = "Region",
                    PivotTableName = "SalesPivot"
                });
                document.AddWorkbookTimelineCache(new ExcelTimelineCacheOptions {
                    Name = "OrderDateTimeline",
                    SourceName = "OrderDate",
                    PivotTableName = "SalesPivot"
                });
                document.AddWorkbookSlicerCache(new ExcelSlicerCacheOptions {
                    Xml = "<customSlicer name=\"CustomRegion\" sourceName=\"Region\" pivotTableName=\"SalesPivot\"><payload /></customSlicer>"
                });
                document.AddWorkbookTimelineCache(new ExcelTimelineCacheOptions {
                    Xml = "<customTimeline name=\"CustomOrderDate\" sourceName=\"OrderDate\" pivotTableName=\"SalesPivot\"><payload /></customTimeline>"
                });
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
                Assert.Equal(0, snapshot.SlicerPartCount);
                Assert.Equal(0, snapshot.TimelinePartCount);
                Assert.False(snapshot.HasSlicers);
                Assert.False(snapshot.HasTimelines);
                Assert.Equal(1, snapshot.SlicerBindingMetadataPartCount);
                Assert.Equal(1, snapshot.TimelineBindingMetadataPartCount);
                Assert.True(snapshot.HasSlicerBindingMetadata);
                Assert.True(snapshot.HasTimelineBindingMetadata);
                Assert.Contains(document.GetWorkbookSlicerCaches(), cache => cache.Name == "CustomRegion");
                Assert.Contains(document.GetWorkbookTimelineCaches(), cache => cache.Name == "CustomOrderDate");

                ExcelFeatureReport report = document.InspectFeatures();
                Assert.Empty(report.FindFeatures("Slicers"));
                Assert.Empty(report.FindFeatures("Timelines"));
                Assert.Equal(OfficeFeatureSupportLevel.Editable, report.FindFeatures("Slicer binding metadata").Single().SupportLevel);
                Assert.Equal(OfficeFeatureSupportLevel.Editable, report.FindFeatures("Timeline binding metadata").Single().SupportLevel);
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                string? metadata = Assert.Single(
                    workbookPart.CustomXmlParts.Select(ReadPivotInteractionMetadataText),
                    text => text != null);
                Assert.NotNull(metadata);
                Assert.Contains("RegionSlicer", metadata);
                Assert.Contains("OrderDateTimeline", metadata);
                Assert.Contains("customSlicer", metadata);
                Assert.Contains("customTimeline", metadata);
            }
        }

        [Fact]
        public void Test_PivotInteractionMetadata_MigratesLegacyCombinedPartOnMutation() {
            string filePath = Path.Combine(_directoryWithFiles, "PackageInteractions.LegacyPivotMetadataMigration.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Region");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WriteExtendedPart(
                    spreadsheet.WorkbookPart!.AddExtendedPart(
                        "https://schemas.evotec.xyz/officeimo/excel/relationships/pivotInteractionMetadata",
                        "application/vnd.officeimo.excel.pivot-interaction-metadata+xml",
                        "xml"),
                    "<pivotInteractionBindings xmlns=\"https://schemas.evotec.xyz/officeimo/excel\">"
                    + "<pivotSlicerBinding name=\"LegacyRegion\" sourceName=\"Region\" pivotTableName=\"SalesPivot\"/>"
                    + "</pivotInteractionBindings>");
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                Assert.Equal("LegacyRegion", Assert.Single(document.GetWorkbookSlicerCaches()).Name);
                document.AddWorkbookTimelineCache(new ExcelTimelineCacheOptions {
                    Name = "CurrentOrderDate",
                    SourceName = "OrderDate",
                    PivotTableName = "SalesPivot"
                });
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                Assert.DoesNotContain(workbookPart.Parts, pair =>
                    string.Equals(pair.OpenXmlPart.ContentType,
                        "application/vnd.officeimo.excel.pivot-interaction-metadata+xml",
                        StringComparison.OrdinalIgnoreCase));
                string metadata = Assert.Single(
                    workbookPart.CustomXmlParts.Select(ReadPivotInteractionMetadataText),
                    text => text != null)!;
                Assert.Contains("LegacyRegion", metadata);
                Assert.Contains("CurrentOrderDate", metadata);
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
                Assert.Equal("LegacyRegion", Assert.Single(document.GetWorkbookSlicerCaches()).Name);
                Assert.Equal("CurrentOrderDate", Assert.Single(document.GetWorkbookTimelineCaches()).Name);
            }
        }

        [Fact]
        public void Test_PivotInteractionInspection_IgnoresOversizedUnrelatedCustomXml() {
            string filePath = Path.Combine(_directoryWithFiles, "PackageInteractions.LargeUnrelatedCustomXml.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Data").CellValue(1, 1, "Value");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                CustomXmlPart part = spreadsheet.WorkbookPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                string xml = "<externalMetadata>" + new string('x', 1_000_100) + "</externalMetadata>";
                using var stream = new MemoryStream(Encoding.UTF8.GetBytes(xml));
                part.FeedData(stream);
            }

            using ExcelDocument inspected = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            Assert.Empty(inspected.GetWorkbookSlicerCaches());
            Assert.Empty(inspected.GetWorkbookTimelineCaches());
            ExcelWorkbookSnapshot snapshot = inspected.CreateInspectionSnapshot();
            Assert.Equal(0, snapshot.SlicerBindingMetadataPartCount);
            Assert.Equal(0, snapshot.TimelineBindingMetadataPartCount);
            ExcelFeatureFinding customXml = Assert.Single(inspected.InspectFeatures().FindFeatures("Custom XML parts"));
            Assert.Equal(1, customXml.Count);
        }

        private static void WriteExtendedPart(ExtendedPart part, string xml) {
            using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
            byte[] bytes = Encoding.UTF8.GetBytes(xml);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static string? ReadPivotInteractionMetadataText(CustomXmlPart part) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            string text = reader.ReadToEnd();
            try {
                XDocument xml = XDocument.Parse(text);
                return xml.Root != null
                    && string.Equals(xml.Root.Name.LocalName, "pivotInteractionBindings", StringComparison.Ordinal)
                    && string.Equals(xml.Root.Name.NamespaceName, "https://schemas.evotec.xyz/officeimo/excel", StringComparison.Ordinal)
                        ? text
                        : null;
            } catch (System.Xml.XmlException) {
                return null;
            }
        }

        private static string ReadSinglePackagePartText(OpenXmlPartContainer container, string contentTypeMarker, bool skipTypedParts = false) {
            var part = Assert.Single(
                container.Parts.Select(relationship => relationship.OpenXmlPart),
                part => (!skipTypedParts || part is ExtendedPart)
                    && part.ContentType.IndexOf(contentTypeMarker, StringComparison.OrdinalIgnoreCase) >= 0);

            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = new StreamReader(stream, Encoding.UTF8);
            return reader.ReadToEnd();
        }

        private static string? GetWorkbookOverrideContentType(string filePath) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            ZipArchiveEntry entry = archive.GetEntry("[Content_Types].xml")
                ?? throw new InvalidOperationException("Workbook package is missing [Content_Types].xml.");

            using Stream stream = entry.Open();
            XDocument document = XDocument.Load(stream);
            XNamespace ns = "http://schemas.openxmlformats.org/package/2006/content-types";
            return document
                .Root?
                .Elements(ns + "Override")
                .FirstOrDefault(element => string.Equals((string?)element.Attribute("PartName"), "/xl/workbook.xml", StringComparison.OrdinalIgnoreCase))
                ?.Attribute("ContentType")
                ?.Value;
        }
    }
}
