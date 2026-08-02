using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_CellValueValidation_PreservesExclusiveInCellImageOnFailure() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.SetInCellImage(1, 1, TinyPng, altText: "Preserved image");

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                sheet.CellValue(1, 1, new string('x', 32_768)));

            Assert.Equal("value", exception.ParamName);
            ExcelInCellImage image = Assert.Single(sheet.GetInCellImages());
            Assert.Equal("A1", image.CellReference);
            Assert.Equal("Preserved image", image.AltText);
            Assert.Equal(TinyPng, image.Bytes);
            AssertRichImageGraphCounts(document, expected: 1);
        }

        [Fact]
        public void Test_AllowedEditRangeValidation_RejectsInvalidXmlBeforeMutation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.Protect();
            sheet.SetAllowedEditRange("Inputs", new[] { "A1" }, securityDescriptor: "existing");
            string originalXml = sheet.WorksheetPart.Worksheet.OuterXml;

            ArgumentException descriptorException = Assert.Throws<ArgumentException>(() =>
                sheet.SetAllowedEditRange("Inputs", new[] { "B2" }, securityDescriptor: "Bad\u0001Descriptor"));
            Assert.Equal("securityDescriptor", descriptorException.ParamName);
            Assert.Equal(originalXml, sheet.WorksheetPart.Worksheet.OuterXml);

            ArgumentException nameException = Assert.Throws<ArgumentException>(() =>
                sheet.SetAllowedEditRange("Bad\u0001Name", new[] { "C3" }));
            Assert.Equal("name", nameException.ParamName);
            ExcelAllowedEditRangeInfo allowed = Assert.Single(sheet.GetAllowedEditRanges());
            Assert.Equal("Inputs", allowed.Name);
            Assert.Equal(new[] { "A1" }, allowed.Ranges);
            Assert.Equal("existing", allowed.SecurityDescriptor);
        }

        [Fact]
        public void Test_FormulaSyntaxTree_ParsesLongPrefixWithoutQuadraticReferenceSearch() {
            string formula = "=" + string.Concat(Enumerable.Repeat("1+", 16_000)) + "A1";
            var stopwatch = Stopwatch.StartNew();

            ExcelFormulaSyntaxTree tree = ExcelFormulaSyntaxTree.Parse(formula);

            stopwatch.Stop();
            ExcelFormulaReferenceSyntax reference = Assert.Single(tree.Nodes.OfType<ExcelFormulaReferenceSyntax>());
            Assert.Equal("A1", reference.Text);
            Assert.Equal(formula, tree.Text);
            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(5), $"Parsing took {stopwatch.Elapsed}.");
        }

        [Fact]
        public void Test_FormulaInspection_MarksCrossSheetCachesDirtyAfterInputMutation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet calculations = document.AddWorksheet("Calculations");
            data.CellValue(1, 1, 2d);
            calculations.CellValue(1, 1, 4d);
            calculations.CellFormula(1, 1, "Data!A1*2");

            ExcelFormulaCellInfo evaluated = Assert.Single(calculations.GetFormulaCells());
            Assert.Equal("4", evaluated.CachedValue);
            Assert.True(evaluated.State.HasFlag(ExcelFormulaState.Evaluated));
            Assert.False(evaluated.State.HasFlag(ExcelFormulaState.Dirty));

            data.CellValue(1, 1, 5d);

            ExcelFormulaCellInfo stale = Assert.Single(calculations.GetFormulaCells());
            Assert.Equal("4", stale.CachedValue);
            Assert.True(stale.State.HasFlag(ExcelFormulaState.Dirty));
            Assert.True(stale.State.HasFlag(ExcelFormulaState.Deferred));
            Assert.False(stale.State.HasFlag(ExcelFormulaState.Evaluated));
        }
    }
}
