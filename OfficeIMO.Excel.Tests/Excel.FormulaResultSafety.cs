using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("SUM(9E307,9E307)", "e", "#NUM!")]
        [InlineData("AVERAGE(9E307,9E307)", "e", "#NUM!")]
        [InlineData("PRODUCT(9E307,9E307)", "e", "#NUM!")]
        [InlineData("SUMSQ(9E307)", "e", "#NUM!")]
        [InlineData("MEDIAN(9E307,9E307)", "n", "9E307")]
        [InlineData("IFERROR(SUM(9E307,9E307),7)", "n", "7")]
        [InlineData("ISNUMBER(SUM(9E307,9E307))", "b", "0")]
        [InlineData("DATE(1900,0,31)", "e", "#NUM!")]
        [InlineData("DATE(1900,-1,1)", "e", "#NUM!")]
        [InlineData("DATE(1900,1,0)", "n", "0")]
        [InlineData("DATE(1900,2,29)", "n", "60")]
        public void FormulaResults_SaveFiniteNumbersOrTypedErrors(string formula, string type, string expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Data");
            sheet.Cell("A1").SetFormula(formula);
            XElement cell = SaveFormulaSafetyCell(document, "A1");
            Assert.Equal(type, (string?)cell.Attribute("t"));
            string actual = cell.Elements().Single(x => x.Name.LocalName == "v").Value;
            if (type == "n") Assert.Equal(double.Parse(expected, CultureInfo.InvariantCulture), double.Parse(actual, CultureInfo.InvariantCulture));
            else Assert.Equal(expected, actual);
        }

        [Theory]
        [InlineData("TRIM(0.1+0.2)", "0.3")]
        [InlineData("LEFT(0.1+0.2,5)", "0.3")]
        [InlineData("CONCAT(1/3)", "0.333333333333333")]
        [InlineData("CONCAT(1E-5)", "0.00001")]
        [InlineData("CONCAT(1E21)", "1E+21")]
        [InlineData("CONCAT(123456789012345678)", "123456789012346000")]
        [InlineData("CONCAT(TRUE(),FALSE())", "TRUEFALSE")]
        [InlineData("CONCAT(\"001.50\")", "001.50")]
        [InlineData("1E-5&\"\"", "0.00001")]
        public void FormulaResults_TextCoercionPreservesExcelNumericAndTextSemantics(string formula, string expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Data");
            sheet.Cell("A1").SetFormula(formula);
            XElement cell = SaveFormulaSafetyCell(document, "A1");
            Assert.Equal("str", (string?)cell.Attribute("t"));
            Assert.Equal(expected, cell.Elements().Single(x => x.Name.LocalName == "v").Value);
        }

        [Theory]
        [InlineData("literal _x0041_", "literal _x005F_x0041_")]
        [InlineData("_x005F_", "_x005F_x005F_")]
        public void FormulaResults_TextCachesPreserveLiteralSpreadsheetEscapeTokens(string text, string serialized) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Data");
            sheet.Cell("A1").SetFormula("CONCAT(\"" + text + "\")");
            XElement cell = SaveFormulaSafetyCell(document, "A1");
            Assert.Equal(serialized, cell.Elements().Single(x => x.Name.LocalName == "v").Value);
            Assert.True(sheet.TryGetCachedFormulaValue(1, 1, out string? cached));
            Assert.Equal(text, cached);
            Assert.Equal(text, Assert.Single(sheet.GetFormulaCells()).CachedValue);
            using var loaded = ExcelDocument.Load(new MemoryStream(document.ToBytes()));
            Assert.Equal(text, loaded.Sheets[0].Cell("A1").GetValue<string>());
            Assert.True(loaded.Sheets[0].TryGetCachedFormulaValue(1, 1, out string? persisted));
            Assert.Equal(text, persisted);
            Assert.Equal(text, Assert.Single(loaded.Sheets[0].GetFormulaCells()).CachedValue);
        }

        [Theory]
        [InlineData("SUM(0.1,0.2)-0.3", "0")]
        [InlineData("-50000+36123.95=-13876.05", "1")]
        [InlineData("IF(-50000+36123.95=-13876.05,1,0)", "1")]
        [InlineData("1.0000000000001=1", "0")]
        public void FormulaResults_ApplyPrecisionAtTheReportedArithmeticBoundary(string formula, string expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Data");
            sheet.Cell("A1").SetFormula(formula);
            XElement cell = SaveFormulaSafetyCell(document, "A1");
            Assert.Equal(expected, cell.Elements().Single(x => x.Name.LocalName == "v").Value);
        }

        [Theory]
        [InlineData("MID(\"😀ab\",2,1)")]
        [InlineData("LEFT(\"😀\",1)")]
        [InlineData("RIGHT(\"😀\",1)")]
        public void FormulaResults_DeferTextSlicesThatCannotBeStoredAsXml(string formula) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Data");
            sheet.Cell("A1").SetFormula(formula);
            XElement cell = SaveFormulaSafetyCell(document, "A1");
            Assert.Equal(formula, cell.Elements().Single(x => x.Name.LocalName == "f").Value);
            Assert.DoesNotContain(cell.Elements(), x => x.Name.LocalName == "v");
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void FormulaResults_ArrayCachesDoNotPublishNonfiniteValues(bool dynamic) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Data");
            sheet.Cell("A1").SetValue(9E307);
            sheet.Cell("A2").SetValue(9E307);
            if (dynamic) sheet.SetDynamicArrayFormula("B1", "SUM(A1:A2)");
            else sheet.SetArrayFormula("B1:B1", "SUM(A1:A2)");
            XElement cell = SaveFormulaSafetyCell(document, "B1");
            Assert.Equal("e", (string?)cell.Attribute("t"));
            Assert.Equal("#NUM!", cell.Elements().Single(x => x.Name.LocalName == "v").Value);
        }

        [Fact]
        public void FormulaResults_KeepIntermediateAndParenthesizedArithmeticPrecision() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Data");
            sheet.Cell("A1").SetFormula("(43.1-43.2)+1");
            sheet.Cell("A2").SetFormula("(SUM(0.1,0.2)-0.3)");
            sheet.Cell("A3").SetFormula("1/7");
            document.Calculate();
            Assert.True(sheet.TryGetCachedFormulaValue(1, 1, out string? intermediate));
            Assert.Equal((43.1d - 43.2d) + 1d, double.Parse(intermediate!, CultureInfo.InvariantCulture));
            Assert.True(sheet.TryGetCachedFormulaValue(2, 1, out string? parenthesized));
            Assert.Equal((0.1d + 0.2d) - 0.3d, double.Parse(parenthesized!, CultureInfo.InvariantCulture));
            Assert.True(sheet.TryGetCachedFormulaValue(3, 1, out string? division));
            Assert.Equal(1d / 7d, double.Parse(division!, CultureInfo.InvariantCulture));
        }

        [Theory]
        [InlineData(0x0001)]
        [InlineData(0xD800)]
        [InlineData(0xFFFF)]
        public void FormulaResults_DeferCustomTextThatCannotBeStoredAsXml(int character) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Data");
            document.Calculation.RegisterCustomFunction("APPTEXT", (_, _) =>
                ExcelFormulaValue.FromText(new string((char)character, 1)));
            sheet.Cell("A1").SetFormula("APPTEXT()");
            XElement cell = SaveFormulaSafetyCell(document, "A1");
            Assert.DoesNotContain(cell.Elements(), x => x.Name.LocalName == "v");
        }

        private static XElement SaveFormulaSafetyCell(ExcelDocument document, string reference) {
            using var output = new MemoryStream();
            document.Save(output, new ExcelSaveOptions { EvaluateFormulasBeforeSave = true });
            using var archive = new ZipArchive(new MemoryStream(output.ToArray()));
            using var reader = new StreamReader(archive.GetEntry("xl/worksheets/sheet1.xml")!.Open());
            return XDocument.Parse(reader.ReadToEnd()).Descendants()
                .Single(x => x.Name.LocalName == "c" && (string?)x.Attribute("r") == reference);
        }
    }
}
