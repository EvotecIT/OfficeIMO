using System;
using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("SUM(A5:T5)", "SUM(A4:T4)")]
        [InlineData("$B$9+C2", "$B$8+C2")]
        [InlineData("SUM($A$2:A9)*2", "SUM($A$2:A8)*2")]
        [InlineData("B3", "#REF!")]
        [InlineData("SUM(A1:A9)", "SUM(A1:A8)")]
        [InlineData("SUM(A3:A9)", "SUM(A3:A8)")]
        [InlineData("SUM(A1:A2)", "SUM(A1:A2)")]
        [InlineData("b9+1", "B8+1")]
        [InlineData("LOG10(B9)", "LOG10(B8)")]
        [InlineData("IF(B9>0,\"B9\",1.5)", "IF(B8>0,\"B9\",1.5)")]
        [InlineData("Other!B9+B9", "Other!B9+B8")]
        public void DeleteRows_moves_references_below_the_deleted_row(string formula, string expected) {
            Assert.Equal(expected, Mutate(formula, sheet => sheet.DeleteRows(3)));
        }

        [Theory]
        [InlineData("SUM(A5:T5)", "SUM(A7:T7)")]
        [InlineData("SUM(A1:A9)", "SUM(A1:A11)")]
        [InlineData("A1+B2", "A1+B2")]
        [InlineData("$C$3*2", "$C$5*2")]
        [InlineData("LOG10(A3)", "LOG10(A5)")]
        public void InsertRows_moves_and_widens_references(string formula, string expected) {
            Assert.Equal(expected, Mutate(formula, sheet => sheet.InsertRows(3, 2)));
        }

        [Theory]
        [InlineData("SUM(A2:T2)", "SUM(A2:V2)")]
        [InlineData("C1+A1", "E1+A1")]
        [InlineData("$B$1", "$D$1")]
        [InlineData("SUM(C:D)", "SUM(E:F)")]
        [InlineData("LOG10(B1)", "LOG10(D1)")]
        public void InsertColumns_moves_and_widens_references(string formula, string expected) {
            Assert.Equal(expected, Mutate(formula, sheet => sheet.InsertColumnsTransactional(2, 2), formulaColumn: 30));
        }

        [Theory]
        [InlineData("SUM(A2:T2)", "SUM(A2:S2)")]
        [InlineData("D1+A1", "C1+A1")]
        [InlineData("B1", "#REF!")]
        public void DeleteColumns_moves_and_shrinks_references(string formula, string expected) {
            Assert.Equal(expected, Mutate(formula, sheet => sheet.DeleteColumnsTransactional(2), formulaColumn: 30));
        }

        private static string Mutate(string formula, Action<ExcelSheet> edit, int formulaColumn = 1) {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                document.AddSheet("Other");
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.Cell(20, formulaColumn).SetFormula(formula);
                edit(sheet);
                document.Save(stream);
            }

            stream.Position = 0;
            using var reopened = ExcelDocument.Load(stream);
            return Assert.Single(reopened.Sheets[1].GetFormulaCells()).Formula;
        }
    }
}
