using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Compatibility_Corpus_PreserveOnlyMetadata_RoundTripsAndBlocksRiskyWorkflows() {
            string filePath = Path.Combine(_directoryWithFiles, "CompatibilityCorpus.PreserveOnlyMetadata.xlsx");
            byte[] customXmlBytes = Encoding.UTF8.GetBytes("<metadata><source>external-system</source><id>INV-2026-001</id></metadata>");
            byte[] connectionBytes = Encoding.UTF8.GetBytes("<connections xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\"><connection id=\"1\" name=\"ExternalSales\" type=\"5\" refreshedVersion=\"7\"/></connections>");
            byte[] queryTableBytes = Encoding.UTF8.GetBytes("<queryTable xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"ExternalSalesQuery\" connectionId=\"1\"/>");

            try {
                ExcelCompatibilityCorpusBuilder.CreateWorkbook(filePath, document => {
                    var sheet = document.AddWorksheet("Imported");
                    sheet.CellValue(1, 1, "Resource");
                    sheet.CellValue(1, 2, "Amount");
                    sheet.SetHyperlink(2, 1, "https://example.org/external-system/invoice/INV-2026-001", display: "Invoice");
                    sheet.CellValue(2, 2, 1250d);
                });

                AddPreserveOnlyPackageParts(filePath, customXmlBytes, connectionBytes, queryTableBytes);

                using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                    document["Imported"].CellValue(3, 2, 1400d);
                    document.Save();
                }

                AssertPreservedPackageParts(filePath, customXmlBytes, connectionBytes, queryTableBytes);

                using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    ExcelFeatureReport report = document.InspectFeatures();

                    Assert.True(report.Can(ExcelPreflightCapability.ReadWorkbookData));
                    Assert.True(report.Can(ExcelPreflightCapability.EditCellValues));
                    Assert.False(report.Can(ExcelPreflightCapability.EditWorkbookStructure));
                    Assert.False(report.Can(ExcelPreflightCapability.BindTemplate));
                    Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                    Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "External hyperlinks"
                        && feature.Details.Any(detail => detail.Contains("INV-2026-001", StringComparison.OrdinalIgnoreCase)));
                    Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Custom XML parts");
                    Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Connections and query tables"
                        && feature.Count == 2);

                    string diagnostics = string.Join(Environment.NewLine,
                        report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport));
                    Assert.Contains("Custom XML parts", diagnostics);
                    Assert.Contains("Connections and query tables", diagnostics);
                    Assert.Contains("Connections and query tables", Assert.Throws<InvalidOperationException>(() =>
                        report.EnsureCan(ExcelPreflightCapability.ExportPdfReport)).Message);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Compatibility_Corpus_MacroEmbeddedObjectsAndControls_RoundTripAndBlockRiskyWorkflows() {
            string filePath = Path.Combine(_directoryWithFiles, "CompatibilityCorpus.MacroEmbeddedControls.xlsx");
            byte[] vbaBytes = Encoding.ASCII.GetBytes("OfficeIMO macro project placeholder");
            byte[] embeddedBytes = Encoding.ASCII.GetBytes("OfficeIMO embedded workbook placeholder");

            try {
                ExcelCompatibilityCorpusBuilder.CreateWorkbook(filePath, document => {
                    var sheet = document.AddWorksheet("Controls");
                    sheet.CellValue(1, 1, "Status");
                    sheet.CellValue(2, 1, "Before");
                });

                AddAdvancedPackageParts(filePath, vbaBytes, embeddedBytes);

                using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                    document["Controls"].CellValue(3, 1, "After");
                    document.Save();
                }

                AssertPreservedAdvancedPackageParts(filePath, vbaBytes, embeddedBytes);

                using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    ExcelFeatureReport report = document.InspectFeatures();

                    Assert.True(report.Can(ExcelPreflightCapability.ReadWorkbookData));
                    Assert.True(report.Can(ExcelPreflightCapability.EditCellValues));
                    Assert.False(report.Can(ExcelPreflightCapability.EditWorkbookStructure));
                    Assert.False(report.Can(ExcelPreflightCapability.BindTemplate));
                    Assert.False(report.Can(ExcelPreflightCapability.ExportPdfReport));

                    Assert.Contains(report.PreservedFeatures, feature => feature.Name == "VBA macros"
                        && feature.Details.Any(detail => detail.Contains("vbaProject", StringComparison.OrdinalIgnoreCase)));
                    Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Embedded packages"
                        && feature.Details.Any(detail => detail.Contains("/embeddings/", StringComparison.OrdinalIgnoreCase)));
                    Assert.Contains(report.PreservedFeatures, feature => feature.Name == "OLE objects"
                        && feature.Details.Any(detail => detail.Contains("Controls", StringComparison.OrdinalIgnoreCase)));
                    Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Form controls"
                        && feature.Details.Any(detail => detail.Contains("Controls", StringComparison.OrdinalIgnoreCase)));

                    string diagnostics = string.Join(Environment.NewLine,
                        report.GetCapabilityDiagnostics(ExcelPreflightCapability.BindTemplate));
                    Assert.Contains("VBA macros", diagnostics);
                    Assert.Contains("Embedded packages", diagnostics);
                    Assert.Contains("OLE objects", diagnostics);
                    Assert.Contains("Form controls", diagnostics);
                    Assert.Contains("VBA macros", Assert.Throws<InvalidOperationException>(() =>
                        report.EnsureCan(ExcelPreflightCapability.BindTemplate)).Message);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static void AddPreserveOnlyPackageParts(
            string filePath,
            byte[] customXmlBytes,
            byte[] connectionBytes,
            byte[] queryTableBytes) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();

            CustomXmlPart customXmlPart = workbookPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            FeedPartData(customXmlPart, customXmlBytes);

            ExtendedPart connectionPart = workbookPart.AddExtendedPart(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml",
                "xml");
            FeedPartData(connectionPart, connectionBytes);

            ExtendedPart queryTablePart = worksheetPart.AddExtendedPart(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
                "xml");
            FeedPartData(queryTablePart, queryTableBytes);
        }

        private static void AddAdvancedPackageParts(
            string filePath,
            byte[] vbaBytes,
            byte[] embeddedBytes) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            VbaProjectPart vbaProjectPart = workbookPart.AddNewPart<VbaProjectPart>();
            FeedPartData(vbaProjectPart, vbaBytes);

            WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
            EmbeddedPackagePart embeddedPackagePart = worksheetPart.AddEmbeddedPackagePart(EmbeddedPackagePartType.Xlsx);
            FeedPartData(embeddedPackagePart, embeddedBytes);
            worksheetPart.Worksheet.Append(new OleObjects(
                "<x:oleObjects xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><x:oleObject progId=\"Package\" shapeId=\"1025\" r:id=\"rIdOlePackage\" /></x:oleObjects>"));
            worksheetPart.Worksheet.Append(new Controls(
                "<x:controls xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><x:control shapeId=\"1026\" name=\"ApproveButton\" r:id=\"rIdControl1\" /></x:controls>"));
            worksheetPart.Worksheet.Save();
        }

        private static void AssertPreservedPackageParts(
            string filePath,
            byte[] customXmlBytes,
            byte[] connectionBytes,
            byte[] queryTableBytes) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
            HyperlinkRelationship hyperlink = Assert.Single(worksheetPart.HyperlinkRelationships);
            Assert.Equal(new Uri("https://example.org/external-system/invoice/INV-2026-001"), hyperlink.Uri);

            CustomXmlPart customXmlPart = Assert.Single(workbookPart.CustomXmlParts);
            Assert.Equal(customXmlBytes, ReadPartBytes(customXmlPart));

            OpenXmlPart connectionPart = Assert.Single(workbookPart.Parts.Select(part => part.OpenXmlPart),
                part => part.ContentType.IndexOf("connections", StringComparison.OrdinalIgnoreCase) >= 0);
            Assert.Equal(connectionBytes, ReadPartBytes(connectionPart));

            OpenXmlPart queryTablePart = Assert.Single(worksheetPart.Parts.Select(part => part.OpenXmlPart),
                part => part.ContentType.IndexOf("queryTable", StringComparison.OrdinalIgnoreCase) >= 0);
            Assert.Equal(queryTableBytes, ReadPartBytes(queryTablePart));
        }

        private static void AssertPreservedAdvancedPackageParts(
            string filePath,
            byte[] vbaBytes,
            byte[] embeddedBytes) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            Assert.NotNull(workbookPart.VbaProjectPart);
            Assert.Equal(vbaBytes, ReadPartBytes(workbookPart.VbaProjectPart!));

            WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
            EmbeddedPackagePart embeddedPackagePart = Assert.Single(worksheetPart.EmbeddedPackageParts);
            Assert.Equal(embeddedBytes, ReadPartBytes(embeddedPackagePart));

            Worksheet worksheet = worksheetPart.Worksheet;
            OleObjects oleObjects = Assert.Single(worksheet.Elements<OleObjects>());
            Controls controls = Assert.Single(worksheet.Elements<Controls>());
            Assert.Contains("rIdOlePackage", oleObjects.OuterXml);
            Assert.Contains("ApproveButton", controls.OuterXml);
        }

        private static void FeedPartData(OpenXmlPart part, byte[] bytes) {
            using var stream = new MemoryStream(bytes);
            part.FeedData(stream);
        }

        private static byte[] ReadPartBytes(OpenXmlPart part) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        }
    }
}
