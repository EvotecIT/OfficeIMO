using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_FormulaReplacement_PreservesCacheButMarksOrdinaryAndArrayFormulasDirty() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Calc");

            sheet.CellValue(1, 1, 4d);
            sheet.CellValue(1, 2, 4d);
            sheet.CellFormula(1, 1, "2+2");
            sheet.CellFormula(1, 2, "2+2");

            Assert.All(sheet.GetFormulaCells(), formula => {
                Assert.Equal("4", formula.CachedValue);
                Assert.True(formula.State.HasFlag(ExcelFormulaState.Evaluated));
                Assert.False(formula.State.HasFlag(ExcelFormulaState.Dirty));
            });

            sheet.CellFormula(1, 1, "3+3");
            sheet.SetArrayFormula("B1:B1", "3+3");

            Assert.All(sheet.GetFormulaCells(), formula => {
                Assert.Equal("4", formula.CachedValue);
                Assert.True(formula.State.HasFlag(ExcelFormulaState.Dirty));
                Assert.True(formula.State.HasFlag(ExcelFormulaState.Deferred));
                Assert.False(formula.State.HasFlag(ExcelFormulaState.Evaluated));
            });
        }

        [Fact]
        public void Test_NamedStyleValidation_RejectsInvalidXmlBeforeStylesheetMutation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetBold();
            string originalStyles = document.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!.OuterXml;

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                sheet.DefineNamedStyle("Bad\u0001Style", 1, 1));

            Assert.Equal("name", exception.ParamName);
            Assert.Equal(originalStyles, document.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!.OuterXml);
            Assert.DoesNotContain(document.GetNamedStyles(), style => style.Name.Contains("Bad", StringComparison.Ordinal));
        }

        [Fact]
        public async Task Test_QueryBackedTable_RejectsDuplicateConnectionInsideTransactionalOwner() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            using var start = new ManualResetEventSlim(false);
            int successes = 0;
            Task[] callers = Enumerable.Range(0, 8).Select(index => Task.Run(() => {
                start.Wait();
                try {
                    document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                        ConnectionName = "SharedConnection",
                        WorksheetName = sheet.Name,
                        StartCell = $"A{index + 1}",
                        TableName = $"Results{index + 1}",
                        ColumnNames = new[] { "Value" }
                    });
                    Interlocked.Increment(ref successes);
                } catch (InvalidOperationException) {
                    // Exactly one caller owns the connection; every contender must be rejected.
                }
            })).ToArray();

            start.Set();
            await Task.WhenAll(callers);

            Assert.Equal(1, successes);
            ExcelQueryBackedTableInfo query = Assert.Single(document.GetQueryBackedTables());
            Assert.StartsWith("Results", query.TableName, StringComparison.Ordinal);
            Assert.Single(sheet.WorksheetPart.TableDefinitionParts);
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
