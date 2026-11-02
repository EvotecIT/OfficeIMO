using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static byte[] CreateDeletionWorkbook() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = document.AddSheet("Data");
                for (int row = 1; row <= 12; row++) {
                    sheet.Cell(row, 1).SetValue(row * 10);
                    sheet.Cell(row, 2).SetFormula("SUM(A1:A" + row + ")+A" + row);
                }
                document.AddSheet("Other").Cell(1, 1).SetValue("Kept");
                document.Save(stream);
            }
            return stream.ToArray();
        }

        public static IEnumerable<object[]> DeletionReadRoutes() {
            yield return new object[] { "TryGetCellText", (Func<ExcelDocument, string>)(document => document.Sheets[0].TryGetCellText(9, 1, out string text) ? text : "none") };
            yield return new object[] { "GetFormulaText", (Func<ExcelDocument, string>)(document => document.Sheets[0].GetFormulaText(9, 2) ?? "none") };
            yield return new object[] { "EnumerateCells", (Func<ExcelDocument, string>)(document => document.Sheets[0].EnumerateCells().Count().ToString()) };
            yield return new object[] { "GetFormulaCells", (Func<ExcelDocument, string>)(document => document.Sheets[0].GetFormulaCells().Last().Formula) };
            yield return new object[] { "CellAt", (Func<ExcelDocument, string>)(document => document.Sheets[0].Cell(9, 1).GetValue<double>().ToString(System.Globalization.CultureInfo.InvariantCulture)) };
            yield return new object[] { "SheetsReopened", (Func<ExcelDocument, string>)(document => document.Sheets.First(sheet => sheet.Name == "Data").TryGetCellText(9, 1, out string text) ? text : "none") };
            yield return new object[] { "OpenXmlDocument", (Func<ExcelDocument, string>)(document => document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Sum(part => part.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>().Count()).ToString()) };
            yield return new object[] { "OtherSheetWrite", (Func<ExcelDocument, string>)(document => {
                document.Sheets[1].Cell(2, 1).SetValue("x");
                return document.Sheets[0].TryGetCellText(9, 1, out string text) ? text : "none";
            }) };
            yield return new object[] { "Calculate", (Func<ExcelDocument, string>)(document => {
                document.Calculate();
                document.Sheets[0].TryGetCachedFormulaValue(9, 2, out string? value);
                return value ?? "none";
            }) };
            yield return new object[] { "Save", (Func<ExcelDocument, string>)(document => {
                using var output = new MemoryStream();
                document.Save(output);
                output.Position = 0;
                using var reopened = ExcelDocument.Load(output);
                return reopened.Sheets[0].TryGetCellText(9, 1, out string text) ? text : "none";
            }) };
        }

        [Theory]
        [MemberData(nameof(DeletionReadRoutes))]
        public void Reads_after_repeated_row_deletes_see_them_as_a_single_call_would(string route, Func<ExcelDocument, string> read) {
            string Run(Action<ExcelSheet> delete) {
                using var document = ExcelDocument.Load(new MemoryStream(CreateDeletionWorkbook()));
                delete(document.Sheets[0]);
                return read(document);
            }

            string repeated = Run(sheet => {
                sheet.DeleteRows(6);
                sheet.DeleteRows(3);
                sheet.DeleteRows(2);
            });
            // Reading between calls flushes each deletion before the next, which is the one-call-at-a-time result.
            string immediate = Run(sheet => {
                _ = sheet.TryGetCellText(1, 1, out _);
                sheet.DeleteRows(6);
                _ = sheet.TryGetCellText(1, 1, out _);
                sheet.DeleteRows(3);
                _ = sheet.TryGetCellText(1, 1, out _);
                sheet.DeleteRows(2);
            });

            Assert.False(string.IsNullOrEmpty(route));
            Assert.Equal(immediate, repeated);
        }

        [Theory]
        [InlineData(new[] { 10, 7, 4 }, "SUM(A1:A9)+A9", "120")]
        [InlineData(new[] { 4, 6, 8 }, "SUM(A1:A9)+A9", "120")]
        [InlineData(new[] { 5, 5, 5, 2 }, "SUM(A1:A8)+A8", "120")]
        public void Deleting_rows_in_any_order_rewrites_the_grid_and_formulas(int[] rows, string lastFormula, string lastValue) {
            using var output = new MemoryStream();
            using (var document = ExcelDocument.Load(new MemoryStream(CreateDeletionWorkbook()))) {
                ExcelSheet sheet = document.Sheets[0];
                foreach (int row in rows) sheet.DeleteRows(row);
                document.Save(output);
            }

            output.Position = 0;
            using (var package = SpreadsheetDocument.Open(output, isEditable: false)) {
                Assert.Empty(new OpenXmlValidator().Validate(package));
            }
            output.Position = 0;
            using var reopened = ExcelDocument.Load(output);
            ExcelSheet saved = reopened.Sheets[0];
            Assert.Equal(12 - rows.Length, saved.GetFormulaCells().Count);
            Assert.Equal(lastFormula, saved.GetFormulaCells().Last().Formula);
            Assert.True(saved.TryGetCellText(12 - rows.Length, 1, out string text));
            Assert.Equal(lastValue, text);
        }

        [Fact]
        public void A_formula_ending_on_a_deleted_row_uses_the_full_reference_rules() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = document.AddSheet("Data");
                for (int row = 1; row <= 6; row++) sheet.Cell(row, 1).SetValue(row);
                sheet.Cell(8, 1).SetFormula("SUM(A2:A5)+A3+A4");
                document.Save(stream);
            }

            using var edited = ExcelDocument.Load(new MemoryStream(stream.ToArray()));
            ExcelSheet target = edited.Sheets[0];
            target.DeleteRows(5);
            target.DeleteRows(3);
            Assert.Equal("SUM(A2:A3)+#REF!+A3", Assert.Single(target.GetFormulaCells()).Formula);
        }
    }
}
