using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_InCellImage_MetadataIsDetachedByClearValueAndFormulaReplacement() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.SetInCellImage(1, 1, TinyPng, altText: "Clear");
            sheet.SetInCellImage(1, 2, TinyPng, altText: "Value");
            sheet.SetInCellImage(1, 3, TinyPng, altText: "Formula");
            sheet.SetInCellImage(1, 4, TinyPng, altText: "Rich text");

            sheet.ClearRange("A1:A1", ExcelClearOptions.Values);
            sheet.CellValue(1, 2, "Replacement");
            sheet.CellFormula(1, 3, "1+1");
            sheet.SetRichText(1, 4, new[] { new ExcelRichTextRun("Replacement") { Bold = true } });

            Assert.Empty(sheet.GetInCellImages());
            Assert.All(sheet.WorksheetPart.Worksheet.Descendants<Cell>(), cell => Assert.Null(cell.ValueMetaIndex));
            Assert.True(sheet.TryGetCellText(1, 2, out string? replacement));
            Assert.Equal("Replacement", replacement);
            Assert.Equal("1+1", Assert.Single(sheet.GetFormulaCells()).Formula);
        }

        [Fact]
        public void Test_PackageWorksheetCopy_StreamsAndRemapsInCellImagesAcrossWorkbooks() {
            using var sourceDocument = ExcelDocument.Create(new MemoryStream());
            ExcelSheet source = sourceDocument.AddWorksheet("Images");
            source.SetInCellImage(2, 2, TinyPng, altText: "Copied image");

            using var targetStream = new MemoryStream();
            using (var targetDocument = ExcelDocument.Create(targetStream)) {
                targetDocument.AddWorksheet("Existing").CellValue(1, 1, "Keep");
                ExcelSheet copied = targetDocument.CopyWorksheetFrom(
                    sourceDocument,
                    "Images",
                    "Copied",
                    SheetNameValidationMode.Sanitize,
                    new ExcelWorksheetCopyOptions { CopyMode = ExcelWorksheetCopyMode.Package });

                ExcelInCellImage image = Assert.Single(copied.GetInCellImages());
                Assert.Equal("B2", image.CellReference);
                Assert.Equal("Copied image", image.AltText);
                Assert.Equal(TinyPng, image.Bytes);
                targetDocument.Save();
            }

            targetStream.Position = 0;
            using var loaded = ExcelDocument.Load(targetStream);
            ExcelInCellImage reloaded = Assert.Single(loaded["Copied"].GetInCellImages());
            Assert.Equal("B2", reloaded.CellReference);
            Assert.Equal("Copied image", reloaded.AltText);
            Assert.Equal(TinyPng, reloaded.Bytes);
            Assert.Empty(loaded.ValidateOpenXml());
        }

        [Fact]
        public void Test_FormulaSyntaxTree_PreservesCompleteErrorLiteralsDuringNameRewrites() {
            const string formula = "=#NULL!+#DIV/0!+#VALUE!+#REF!+#NAME?+#NUM!+#N/A+#GETTING_DATA+#SPILL!+#CALC!+#FIELD!+#BLOCKED!+#UNKNOWN!+#BUSY!+#CONNECT!+#PYTHON!+TaxRate+Sales[Amount]";
            ExcelFormulaSyntaxTree tree = ExcelFormulaSyntaxTree.Parse(formula);

            string names = tree.RewriteNames(name =>
                string.Equals(name, "TaxRate", StringComparison.OrdinalIgnoreCase) ? "Rate2026" : "Unexpected");
            string tables = tree.RewriteTableNames(name =>
                string.Equals(name, "Sales", StringComparison.OrdinalIgnoreCase) ? "Orders" : name);

            Assert.Equal(formula.Replace("TaxRate", "Rate2026", StringComparison.Ordinal), names);
            Assert.Equal(formula.Replace("Sales", "Orders", StringComparison.Ordinal), tables);
            Assert.DoesNotContain(tree.Nodes.OfType<ExcelFormulaNameSyntax>(), node => node.Name.StartsWith("REF", StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(tree.Nodes.OfType<ExcelFormulaNameSyntax>(), node => node.Name.StartsWith("NAME", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void Test_UnicodeSheetQualifier_IsSearchableAndRewrittenByStructuralEdits() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Данные");
            ExcelSheet summary = document.AddWorksheet("Итог");
            data.CellValue(1, 1, 10);
            summary.CellFormula(1, 1, "Данные!A1");

            ExcelFormulaCellInfo match = Assert.Single(document.SearchFormulas(
                new ExcelFormulaSearchOptions { Reference = "Данные!A1" }));
            Assert.Equal("Итог", match.SheetName);

            data.InsertColumns(1);

            Assert.Equal("Данные!B1", Assert.Single(summary.GetFormulaCells()).Formula);
        }
    }
}
