using System.Data;
using System.Text;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelAllSeverityBatch14SecurityTests {
    [Fact]
    public void DirectDataSetFastSaveHonorsSignedWorkbookMutationPolicy() {
        var table = new DataTable("Data");
        table.Columns.Add("Value", typeof(string));
        table.Rows.Add("safe");
        var dataSet = new DataSet("Export");
        dataSet.Tables.Add(table);

        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        document.InsertDataSet(dataSet, autoFit: false);
        document._spreadSheetDocument.AddDigitalSignatureOriginPart();

        Assert.Throws<ExcelSignedWorkbookMutationException>(() => document.ToBytes());
    }

    [Fact]
    public void CustomFormulaDoesNotInvokeCallbackForNonFiniteArguments() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        int callbackCount = 0;
        document.Calculation.RegisterCustomFunction("ECHOFINITE", (_, arguments) => {
            callbackCount++;
            return arguments[0];
        });
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellFormula(1, 1, "ECHOFINITE(1E309)");

        Exception? exception = Record.Exception(() => document.Calculate());

        Assert.Null(exception);
        Assert.Equal(0, callbackCount);
    }

    [Fact]
    public void HeaderFooterPathFieldsRequireExplicitExportOptIn() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-header-path-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string workbookPath = Path.Combine(root, "report.xlsx");
        try {
            using ExcelDocument document = ExcelDocument.Create(workbookPath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            sheet.CellValue(1, 1, "body");
            sheet.SetHeaderFooter(headerLeft: "Path &[Path] &Z");

            string defaultSvg = Encoding.UTF8.GetString(Assert.Single(sheet.ExportImages(
                OfficeImageExportFormat.Svg,
                new ExcelWorksheetImageExportOptions {
                    Range = "A1:AZ2",
                    SplitByManualPageBreaks = true
                })).Bytes);
            string optedInSvg = Encoding.UTF8.GetString(Assert.Single(sheet.ExportImages(
                OfficeImageExportFormat.Svg,
                new ExcelWorksheetImageExportOptions {
                    Range = "A1:AZ2",
                    SplitByManualPageBreaks = true,
                    IncludeWorkbookPathInHeaderFooter = true
                })).Bytes);

            Assert.DoesNotContain(root, defaultSvg, StringComparison.Ordinal);
            Assert.Contains(root, optedInSvg, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }
    }

    [Fact]
    public void LoadSkipsMalformedCustomPropertiesAndRetainsValidOnes() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-custom-property-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                document.SetCustomDocumentProperty("Valid", "retained");
                document.AddWorksheet("Data").CellValue(1, 1, "body");
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                CustomFilePropertiesPart part = package.CustomFilePropertiesPart!;
                part.Properties!.Append(new CustomDocumentProperty(new VTInt32("not-an-integer")) {
                    FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                    PropertyId = 99,
                    Name = "Malformed"
                });
                part.Properties.Save();
            }

            using ExcelDocument loaded = ExcelDocument.Load(path);
            Assert.Equal("retained", loaded.CustomDocumentProperties["Valid"].Text);
            Assert.False(loaded.CustomDocumentProperties.ContainsKey("Malformed"));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
