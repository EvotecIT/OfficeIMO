using System;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_TransactionRollback_RestoresExclusiveInCellImageGraph() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.SetInCellImage(1, 1, TinyPng, altText: "Rollback image");
            AssertRichImageGraphCounts(document, expected: 1);
            using var cancellation = new CancellationTokenSource();

            Assert.Throws<OperationCanceledException>(() => sheet.ApplyTransactionalMutation(_ => {
                Assert.True(sheet.RemoveInCellImage(1, 1));
                Assert.Empty(sheet.GetInCellImages());
                Assert.Empty(document.WorkbookPartRoot.RdRichValueParts);
                cancellation.Cancel();
            }, 0, new ExcelMutationPlanOptions(), cancellation.Token));

            ExcelInCellImage restored = Assert.Single(sheet.GetInCellImages());
            Assert.Equal("A1", restored.CellReference);
            Assert.Equal("Rollback image", restored.AltText);
            Assert.Equal(TinyPng, restored.Bytes);
            AssertRichImageGraphCounts(document, expected: 1);
            Assert.Empty(document.ValidateOpenXml());

            sheet.SetInCellImage(1, 2, TinyPng, altText: "Second image");
            using var renumberCancellation = new CancellationTokenSource();
            Assert.Throws<OperationCanceledException>(() => sheet.ApplyTransactionalMutation(_ => {
                Assert.True(sheet.RemoveInCellImage(1, 1));
                Assert.Equal("B1", Assert.Single(sheet.GetInCellImages()).CellReference);
                renumberCancellation.Cancel();
            }, 0, new ExcelMutationPlanOptions(), renumberCancellation.Token));
            Assert.Equal(new[] { "A1", "B1" }, sheet.GetInCellImages()
                .Select(image => image.CellReference)
                .OrderBy(reference => reference, StringComparer.Ordinal)
                .ToArray());
            AssertRichImageGraphCounts(document, expected: 2);

            using var saved = new MemoryStream();
            document.Save(saved);
            saved.Position = 0;
            using ExcelDocument reloaded = ExcelDocument.Load(saved);
            Assert.All(reloaded["Images"].GetInCellImages(), image => Assert.Equal(TinyPng, image.Bytes));
        }

        [Fact]
        public void Test_TableSchema_RewritesUnqualifiedStructuredReferencesOnlyInsideOwnerTable() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Amount");
            sheet.CellValue(1, 2, "Tax");
            sheet.CellFormula(2, 1, "=[@Amount]*2");
            sheet.CellFormula(3, 1, "=[@Tax]+1");
            sheet.CellValue(2, 2, 1);
            sheet.CellValue(3, 2, 2);
            sheet.CellFormula(1, 4, "=[@Amount]");
            sheet.AddTable("A1:B3", true, "Sales", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
            string? FormulaAt(string reference) => sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => string.Equals(cell.CellReference?.Value, reference, StringComparison.Ordinal))
                .CellFormula?.Text;

            sheet.SetTableSchema("Sales", new[] { "Net", "Fee" });

            Assert.Equal("[@Net]*2", FormulaAt("A2"));
            Assert.Equal("[@Fee]+1", FormulaAt("A3"));
            Assert.Equal("[@Amount]", FormulaAt("D1"));

            sheet.SetTableSchema("Sales", new[] { "Net" }, "A1:A3");

            Assert.Equal("[@Net]*2", FormulaAt("A2"));
            Assert.Equal("#REF!+1", FormulaAt("A3"));
            Assert.Equal("[@Amount]", FormulaAt("D1"));
        }

        [Fact]
        public void Test_FeatureReport_RegistersInCellImagesWithoutReadingPayloads() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.SetInCellImage(2, 2, TinyPng, altText: "Native image");

            ExcelFeatureReport report = document.InspectFeatures();

            Assert.Equal(1, Assert.Single(report.FindFeatures("Images")).Count);
            ExcelFeatureFinding unsupported = Assert.Single(report.FindFeatures("PDF-unsupported images"));
            Assert.Equal(1, unsupported.Count);
            Assert.Contains(unsupported.Details, detail =>
                detail.Contains("Images!B2", StringComparison.Ordinal)
                && detail.Contains("in-cell", StringComparison.OrdinalIgnoreCase));
            Assert.False(report.CanExportPdfReport);
        }
    }
}
