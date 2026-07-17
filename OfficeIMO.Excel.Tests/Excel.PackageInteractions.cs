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

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
                    "<pivotSlicerBinding xmlns=\"https://schemas.evotec.xyz/officeimo/excel\" name=\"LegacyRegion\" sourceName=\"Region\" pivotTableName=\"SalesPivot\"/>");
                WriteExtendedPart(
                    workbookPart.AddExtendedPart(
                        "http://schemas.microsoft.com/office/2007/relationships/slicerCache",
                        "application/vnd.ms-excel.slicerCache+xml",
                        "xml"),
                    "<slicerCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" name=\"NativeRegion\"/>");
                WriteExtendedPart(
                    workbookPart.AddExtendedPart(
                        "http://schemas.microsoft.com/office/2011/relationships/timelineCache",
                        "application/vnd.ms-excel.timelineCache+xml",
                        "xml"),
                    "<pivotTimelineBinding xmlns=\"https://schemas.evotec.xyz/officeimo/excel\" name=\"LegacyOrderDate\" sourceName=\"OrderDate\" pivotTableName=\"SalesPivot\"/>");
                WriteExtendedPart(
                    workbookPart.AddExtendedPart(
                        "http://schemas.microsoft.com/office/2011/relationships/timelineCache",
                        "application/vnd.ms-excel.timelineCache+xml",
                        "xml"),
                    "<timelineCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2011/1/timeline\" name=\"NativeOrderDate\"/>");
            }

            using (var document = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
                ExcelPivotInteractionCacheInfo slicer = Assert.Single(document.GetWorkbookSlicerCaches());
                Assert.Equal("LegacyRegion", slicer.Name);
                Assert.Equal("Region", slicer.SourceName);
                Assert.Equal("SalesPivot", slicer.PivotTableName);

                ExcelPivotInteractionCacheInfo timeline = Assert.Single(document.GetWorkbookTimelineCaches());
                Assert.Equal("LegacyOrderDate", timeline.Name);
                Assert.Equal("OrderDate", timeline.SourceName);
                Assert.Equal("SalesPivot", timeline.PivotTableName);

                ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
                Assert.Equal(1, snapshot.SlicerPartCount);
                Assert.Equal(1, snapshot.TimelinePartCount);
                Assert.Equal(1, snapshot.SlicerBindingMetadataPartCount);
                Assert.Equal(1, snapshot.TimelineBindingMetadataPartCount);
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

            using (var document = ExcelDocument.Load(destinationPath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
                OpenXmlPart part = document.AddWorkbookConnectionMetadata(connectionXml);
                Assert.IsType<ConnectionsPart>(part);
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
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
                Assert.Equal(0, snapshot.SlicerPartCount);
                Assert.Equal(0, snapshot.TimelinePartCount);
                Assert.False(snapshot.HasSlicers);
                Assert.False(snapshot.HasTimelines);
                Assert.Equal(1, snapshot.SlicerBindingMetadataPartCount);
                Assert.Equal(1, snapshot.TimelineBindingMetadataPartCount);
                Assert.True(snapshot.HasSlicerBindingMetadata);
                Assert.True(snapshot.HasTimelineBindingMetadata);

                ExcelFeatureReport report = document.InspectFeatures();
                Assert.Empty(report.FindFeatures("Slicers"));
                Assert.Empty(report.FindFeatures("Timelines"));
                Assert.Equal(ExcelFeatureSupportLevel.Editable, report.FindFeatures("Slicer binding metadata").Single().SupportLevel);
                Assert.Equal(ExcelFeatureSupportLevel.Editable, report.FindFeatures("Timeline binding metadata").Single().SupportLevel);
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                Assert.Contains("RegionSlicer", ReadSinglePackagePartText(workbookPart, "slicerCache"));
                Assert.Contains("OrderDateTimeline", ReadSinglePackagePartText(workbookPart, "timelineCache"));
            }
        }

        private static void WriteExtendedPart(ExtendedPart part, string xml) {
            using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
            byte[] bytes = Encoding.UTF8.GetBytes(xml);
            stream.Write(bytes, 0, bytes.Length);
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
