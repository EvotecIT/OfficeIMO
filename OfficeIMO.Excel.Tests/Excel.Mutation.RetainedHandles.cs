using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Mutation_rollback_preserves_retained_worksheet_metadata(bool sheetCaching) {
            using var document = ExcelDocument.Create();
            document.SheetCachingEnabled = sheetCaching;
            ExcelSheet sheet = document.AddSheet("Report");
            document.AddSheet("Keep");
            sheet.CellValue(1, 1, "Before rollback");
            Assert.Throws<InvalidOperationException>(() => sheet.ApplyTransactionalMutation(_ => {
                sheet.CellValue(1, 1, "Interrupted edit");
                throw new InvalidOperationException("Abort transaction");
            }, 1, new ExcelMutationPlanOptions(), default));

            sheet.SetHidden(true);
            sheet.Name = "Recovered";
            using var stream = new MemoryStream();
            document.Save(stream);

            using var saved = ExcelDocument.Load(stream);
            Assert.True(saved["Recovered"].Hidden);
            Assert.Equal("Recovered", sheet.Name);
            Assert.Equal("Before rollback", saved["Recovered"].CellAt(1, 1).GetValue<string>());
        }

        [Fact]
        public void Mutation_rollback_of_an_unchanged_loaded_workbook_keeps_handles_and_formula_state_working() {
            using var source = new MemoryStream();
            using (var authored = ExcelDocument.Create()) {
                ExcelSheet report = authored.AddSheet("Report");
                report.CellValue(1, 1, 2d);
                report.CellFormula(1, 2, "A1*2");
                authored.AddSheet("Keep").CellValue(1, 1, "Untouched");
                authored.Save(source);
            }

            using var document = ExcelDocument.Load(new MemoryStream(source.ToArray()));
            ExcelSheet sheet = document.Sheets[0];
            ExcelSheet keep = document.Sheets[1];
            Assert.Throws<InvalidOperationException>(() => sheet.ApplyTransactionalMutation(_ => {
                sheet.CellValue(1, 1, 99d);
                keep.CellValue(1, 1, "Interrupted edit");
                throw new InvalidOperationException("Abort transaction");
            }, 1, new ExcelMutationPlanOptions(), default));

            Assert.Equal(2d, sheet.CellAt(1, 1).GetValue<double>());
            Assert.Equal("Untouched", keep.CellAt(1, 1).GetValue<string>());
            sheet.CellValue(1, 1, 5d);
            document.Calculate();
            using var stream = new MemoryStream();
            document.Save(stream);
            using var saved = ExcelDocument.Load(stream);
            Assert.Equal(5d, saved["Report"].CellAt(1, 1).GetValue<double>());
            Assert.True(saved["Report"].TryGetCachedFormulaValue(1, 2, out string? cached));
            Assert.Equal("10", cached);
        }

        [Fact]
        public void Mutation_rollback_refreshes_retained_alias_cell_caches() {
            using var document = ExcelDocument.Create();
            document.SheetCachingEnabled = false;
            ExcelSheet sheet = document.AddSheet("Report");
            sheet.CellValue(1, 1, "Before rollback");
            ExcelSheet alias = document.Sheets[0];
            Assert.Equal("Before rollback", alias.CellAt(1, 1).GetValue<string>());
            Assert.Throws<InvalidOperationException>(() => sheet.ApplyTransactionalMutation(_ => {
                sheet.CellValue(1, 1, "Interrupted edit");
                throw new InvalidOperationException("Abort transaction");
            }, 1, new ExcelMutationPlanOptions(), default));

            Assert.Equal("Before rollback", alias.CellAt(1, 1).GetValue<string>());
            alias.CellValue(2, 1, "After rollback");
            using var stream = new MemoryStream();
            document.Save(stream);
            using var saved = ExcelDocument.Load(stream);
            Assert.Equal("Before rollback", saved["Report"].CellAt(1, 1).GetValue<string>());
            Assert.Equal("After rollback", saved["Report"].CellAt(2, 1).GetValue<string>());
        }
    }
}
