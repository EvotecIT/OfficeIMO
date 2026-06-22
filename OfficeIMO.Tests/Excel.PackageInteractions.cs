using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_InspectionCountsSlicerAndTimelinePackageParts() {
            string filePath = Path.Combine(_directoryWithFiles, "PackageInteractions.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data").CellValue(1, 1, "Value");
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

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
                Assert.Equal(1, snapshot.SlicerPartCount);
                Assert.Equal(1, snapshot.TimelinePartCount);
                Assert.True(snapshot.HasSlicers);
                Assert.True(snapshot.HasTimelines);
            }
        }

        [Fact]
        public void Test_CopyPackage_PreservesPartsAndNormalizesWorkbookContentType() {
            string sourcePath = Path.Combine(_directoryWithFiles, "PackageClone.Source.xlsx");
            string destinationPath = Path.Combine(_directoryWithFiles, "PackageClone.Target.xlsm");

            using (var document = ExcelDocument.Create(sourcePath)) {
                document.AddWorkSheet("Data").CellValue(1, 1, "Value");
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

            using (var document = ExcelDocument.Load(destinationPath, readOnly: true)) {
                var worksheetPart = document._spreadSheetDocument.WorkbookPart!.WorksheetParts.Single();
                Assert.Equal("Value", GetCellValue(document._spreadSheetDocument, worksheetPart, "A1"));

                ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
                Assert.Equal(1, snapshot.SlicerPartCount);
                Assert.Equal(1, snapshot.TimelinePartCount);
            }
        }

        [Fact]
        public void Test_ConnectionAndQueryTableMetadataParts_AreAuthoredAndInspected() {
            string filePath = Path.Combine(_directoryWithFiles, "PackageInteractions.ConnectionMetadata.xlsx");
            const string connectionXml = "<connections xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\"><connection id=\"1\" name=\"SalesConnection\" type=\"5\" refreshedVersion=\"7\"/></connections>";
            const string queryTableXml = "<queryTable xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"SalesQuery\" connectionId=\"1\"/>";

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data").CellValue(1, 1, "Region");
                document.AddWorkbookConnectionMetadata(connectionXml);
                document.AddWorksheetQueryTableMetadata("Data", queryTableXml);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
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
        public void Test_SlicerAndTimelineMetadataParts_AreAuthoredAndInspected() {
            string filePath = Path.Combine(_directoryWithFiles, "PackageInteractions.SlicerTimelineMetadata.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data").CellValue(1, 1, "Region");
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

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
                Assert.Equal(1, snapshot.SlicerPartCount);
                Assert.Equal(1, snapshot.TimelinePartCount);
                Assert.True(snapshot.HasSlicers);
                Assert.True(snapshot.HasTimelines);

                ExcelFeatureReport report = document.InspectFeatures();
                Assert.Equal(ExcelFeatureSupportLevel.PartiallyEditable, report.FindFeatures("Slicers").Single().SupportLevel);
                Assert.Equal(ExcelFeatureSupportLevel.PartiallyEditable, report.FindFeatures("Timelines").Single().SupportLevel);
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
