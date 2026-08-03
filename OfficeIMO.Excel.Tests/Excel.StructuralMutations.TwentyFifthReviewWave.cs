using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public async Task Test_QueryBackedTable_RefreshPreservesHeaderlessLayout() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            ExcelQueryBackedTableInfo source = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "HeaderlessQuery",
                WorksheetName = sheet.Name,
                TableName = "HeaderlessResults",
                ColumnNames = new[] { "Region", "Amount" }
            });
            Table table = sheet.WorksheetPart.TableDefinitionParts.Single().Table!;
            table.HeaderRowCount = 0U;
            table.GetFirstChild<AutoFilter>()?.Remove();
            table.Save();

            var host = new StubQueryHost(new ExcelQueryExecutionResult(
                new[] { "Region", "Amount" },
                new IReadOnlyList<object?>[] {
                    new object?[] { "East", 10d },
                    new object?[] { "West", 20d }
                }));
            ExcelQueryRefreshResult refreshed = await document.RefreshQueryAsync(
                source.TableName,
                host,
                new ExcelQueryExecutionPolicy { AllowExecution = true });

            Assert.Equal("A1:B2", refreshed.Range);
            Assert.Equal(0U, table.HeaderRowCount!.Value);
            Assert.Null(table.GetFirstChild<AutoFilter>());
            Assert.Equal("East", sheet.CellAt(1, 1).GetValue<string>());
            Assert.Equal(20d, sheet.CellAt(2, 2).GetValue<double>());
            Assert.DoesNotContain(sheet.WorksheetPart.Worksheet.Descendants<Cell>(), cell =>
                cell.CellReference?.Value == "A3");
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_FormulaSearch_SkipsSeparatelyQuotedThreeDimensionalQualifiers() {
            using var document = ExcelDocument.Create(new MemoryStream());
            document.AddWorksheet("SUM(A1)");
            document.AddWorksheet("Last Sheet");
            ExcelSheet summary = document.AddWorksheet("Summary");
            summary.CellFormula(1, 1, "'SUM(A1)':'Last Sheet'!B1");
            summary.CellFormula(2, 1, "SUM(B1)");

            ExcelFormulaCellInfo match = Assert.Single(summary.SearchFormulas(
                new ExcelFormulaSearchOptions { Function = "SUM" }));

            Assert.Equal("A2", match.CellReference);
        }

        [Fact]
        public void Test_FileBackedEdit_UsesOwnerOnlyStagingFileOnUnix() {
#if NET6_0_OR_GREATER
            if (OperatingSystem.IsWindows()) return;
            string path = Path.Combine(_directoryWithFiles, "FileBackedPrivateStaging.xlsx");
            using (var created = ExcelDocument.Create()) {
                created.AddWorksheet("Data").CellValue(1, 1, "private");
                created.Save(path);
            }

            using ExcelDocument document = ExcelDocument.OpenFileBacked(path);
            FieldInfo field = typeof(ExcelDocument).GetField(
                "_ownedOpenStream",
                BindingFlags.Instance | BindingFlags.NonPublic)!;
            FileStream staging = Assert.IsType<FileStream>(field.GetValue(document));
            const UnixFileMode accessBits = UnixFileMode.UserRead
                | UnixFileMode.UserWrite
                | UnixFileMode.UserExecute
                | UnixFileMode.GroupRead
                | UnixFileMode.GroupWrite
                | UnixFileMode.GroupExecute
                | UnixFileMode.OtherRead
                | UnixFileMode.OtherWrite
                | UnixFileMode.OtherExecute;

            Assert.Equal(
                UnixFileMode.UserRead | UnixFileMode.UserWrite,
                File.GetUnixFileMode(staging.Name) & accessBits);
#endif
        }
    }
}
