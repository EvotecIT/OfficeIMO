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
            sheet.Cell(1, 1).SetValue("Before rollback");
            Assert.Throws<InvalidOperationException>(() => sheet.ApplyTransactionalMutation(_ => {
                sheet.Cell(1, 1).SetValue("Interrupted edit");
                throw new InvalidOperationException("Abort transaction");
            }, 1, new ExcelMutationPlanOptions(), default));

            sheet.SetHidden(true);
            sheet.Name = "Recovered";
            using var stream = new MemoryStream();
            document.Save(stream);

            using var saved = ExcelDocument.Load(stream);
            Assert.True(saved["Recovered"].Hidden);
            Assert.Equal("Recovered", sheet.Name);
            Assert.Equal("Before rollback", saved["Recovered"].Cell(1, 1).GetValue<string>());
        }

        [Fact]
        public void Mutation_rollback_of_an_unchanged_loaded_workbook_keeps_handles_and_formula_state_working() {
            using var source = new MemoryStream();
            using (var authored = ExcelDocument.Create()) {
                ExcelSheet report = authored.AddSheet("Report");
                report.Cell(1, 1).SetValue(2d);
                report.Cell(1, 2).SetFormula("A1*2");
                authored.AddSheet("Keep").Cell(1, 1).SetValue("Untouched");
                authored.Save(source);
            }

            using var document = ExcelDocument.Load(new MemoryStream(source.ToArray()));
            ExcelSheet sheet = document.Sheets[0];
            ExcelSheet keep = document.Sheets[1];
            Assert.Throws<InvalidOperationException>(() => sheet.ApplyTransactionalMutation(_ => {
                sheet.Cell(1, 1).SetValue(99d);
                keep.Cell(1, 1).SetValue("Interrupted edit");
                throw new InvalidOperationException("Abort transaction");
            }, 1, new ExcelMutationPlanOptions(), default));

            Assert.Equal(2d, sheet.Cell(1, 1).GetValue<double>());
            Assert.Equal("Untouched", keep.Cell(1, 1).GetValue<string>());
            sheet.Cell(1, 1).SetValue(5d);
            document.Calculate();
            using var stream = new MemoryStream();
            document.Save(stream);
            using var saved = ExcelDocument.Load(stream);
            Assert.Equal(5d, saved["Report"].Cell(1, 1).GetValue<double>());
            Assert.True(saved["Report"].TryGetCachedFormulaValue(1, 2, out string? cached));
            Assert.Equal("10", cached);
        }

        [Fact]
        public void Mutation_rollback_refreshes_retained_alias_cell_caches() {
            using var document = ExcelDocument.Create();
            document.SheetCachingEnabled = false;
            ExcelSheet sheet = document.AddSheet("Report");
            sheet.Cell(1, 1).SetValue("Before rollback");
            ExcelSheet alias = document.Sheets[0];
            Assert.Equal("Before rollback", alias.Cell(1, 1).GetValue<string>());
            Assert.Throws<InvalidOperationException>(() => sheet.ApplyTransactionalMutation(_ => {
                sheet.Cell(1, 1).SetValue("Interrupted edit");
                throw new InvalidOperationException("Abort transaction");
            }, 1, new ExcelMutationPlanOptions(), default));

            Assert.Equal("Before rollback", alias.Cell(1, 1).GetValue<string>());
            alias.Cell(2, 1).SetValue("After rollback");
            using var stream = new MemoryStream();
            document.Save(stream);
            using var saved = ExcelDocument.Load(stream);
            Assert.Equal("Before rollback", saved["Report"].Cell(1, 1).GetValue<string>());
            Assert.Equal("After rollback", saved["Report"].Cell(2, 1).GetValue<string>());
        }
    }
}
