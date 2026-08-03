using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_TransactionRollback_InvalidatesSharedStringAndStyleCaches() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            using var cancellation = new CancellationTokenSource();

            Assert.ThrowsAny<OperationCanceledException>(() => sheet.ApplyTransactionalMutation(token => {
                sheet.CellValue(1, 1, "rollback-only");
                sheet.CellValue(1, 2, new DateTime(2026, 8, 1));
                cancellation.Cancel();
                token.ThrowIfCancellationRequested();
                return 2;
            }, new ExcelMutationPlanOptions(), cancellation.Token));

            sheet.CellValue(1, 1, "rollback-only");
            sheet.CellValue(1, 2, new DateTime(2026, 8, 2));
            Assert.Empty(document.ValidateOpenXml());

            using var stream = new MemoryStream();
            document.Save(stream);
            stream.Position = 0;
            using ExcelDocument loaded = ExcelDocument.Load(stream);
            Assert.Equal("rollback-only", loaded["Data"].CellAt(1, 1).GetValue<string>());
            Assert.Equal(new DateTime(2026, 8, 2),
                DateTime.FromOADate(loaded["Data"].CellAt(1, 2).GetValue<double>()));
        }

        [Fact]
        public void Test_QueryBackedTable_CreationFailuresAreTransactional() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Existing");
            sheet.CellValue(2, 1, "keep");
            sheet.AddTable("A1:A2", true, "ExistingTable", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
            sheet.CellValue(1, 3, "preserve");

            Assert.Throws<InvalidOperationException>(() => document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "OverlapQuery",
                WorksheetName = sheet.Name,
                StartCell = "A1",
                TableName = "OverlapResults",
                ColumnNames = new[] { "Changed" }
            }));
            Assert.Throws<ArgumentException>(() => document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "InvalidNameQuery",
                WorksheetName = sheet.Name,
                StartCell = "C1",
                TableName = "invalid table name",
                ColumnNames = new[] { "Changed" }
            }));

            Assert.Equal("Existing", sheet.CellAt(1, 1).GetValue<string>());
            Assert.Equal("preserve", sheet.CellAt(1, 3).GetValue<string>());
            Assert.Single(sheet.WorksheetPart.TableDefinitionParts);
            Assert.Empty(document.GetQueryBackedTables());
            Assert.Null(document.WorkbookPartRoot.ConnectionsPart);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_FormulaSearch_SkipsStructuredReferenceTokens() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellFormula(1, 1, "Table1[SUM(]");
            sheet.CellFormula(2, 1, "SUM(B1)");
            sheet.CellFormula(3, 1, "Table1[[#Headers],[SUM(]]]");

            ExcelFormulaCellInfo match = Assert.Single(sheet.SearchFormulas(
                new ExcelFormulaSearchOptions { Function = "SUM" }));

            Assert.Equal("A2", match.CellReference);
        }

        [Fact]
        public void Test_RemoveWorksheet_PrunesUnusedOwnedQueryConnection() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            document.AddWorksheet("Keep");
            ExcelQueryBackedTableInfo query = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "DisposableQuery",
                CommandText = "sensitive-command",
                WorksheetName = sheet.Name,
                TableName = "DisposableResults",
                ColumnNames = new[] { "Value" },
                RefreshOnOpen = true
            });
            Assert.Contains(document.WorkbookPartRoot.ConnectionsPart!.Connections!.Elements<Connection>(),
                connection => connection.Id?.Value == query.ConnectionId);

            document.RemoveWorksheet(sheet.Name);

            Assert.Empty(document.GetQueryBackedTables());
            Assert.True(document.WorkbookPartRoot.ConnectionsPart == null
                || !document.WorkbookPartRoot.ConnectionsPart.Connections!.Elements<Connection>()
                    .Any(connection => connection.Id?.Value == query.ConnectionId));
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_RangeCriteria_TargetTableOwnedAutoFilter() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "A");
            sheet.CellValue(2, 2, 1d);
            sheet.CellValue(3, 1, string.Empty);
            sheet.CellValue(3, 2, 2d);
            sheet.AddTable("A1:B3", true, "FilteredTable", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
            ExcelAutoFilterInfo filterInfo = Assert.Single(sheet.GetAutoFilters(), filter => filter.IsTableFilter);

            sheet.AutoFilterBlanks(filterInfo.Range, 0U);
            sheet.AutoFilterTopBottom(filterInfo.Range, 1U, 1);

            Assert.Null(sheet.WorksheetPart.Worksheet.GetFirstChild<AutoFilter>());
            Table table = sheet.WorksheetPart.TableDefinitionParts.Single().Table!;
            AutoFilter tableFilter = table.GetFirstChild<AutoFilter>()!;
            Assert.True(tableFilter.Elements<FilterColumn>().Single(column => column.ColumnId?.Value == 0U)
                .GetFirstChild<Filters>()!.Blank!.Value);
            Assert.Equal(1d, tableFilter.Elements<FilterColumn>().Single(column => column.ColumnId?.Value == 1U)
                .GetFirstChild<Top10>()!.Val!.Value);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public async Task Test_QueryBackedTable_RejectsNonFiniteNumbersBeforeMutation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            ExcelQueryBackedTableInfo query = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "FiniteQuery",
                WorksheetName = sheet.Name,
                TableName = "FiniteResults",
                ColumnNames = new[] { "Value" }
            });

            foreach (object value in new object[] { double.NaN, double.PositiveInfinity, float.NegativeInfinity }) {
                var host = new StubQueryHost(new ExcelQueryExecutionResult(
                    new[] { "Value" },
                    new IReadOnlyList<object?>[] { new object?[] { value } }));
                await Assert.ThrowsAsync<InvalidDataException>(() => document.RefreshQueryAsync(
                    query.TableName,
                    host,
                    new ExcelQueryExecutionPolicy { AllowExecution = true }));
            }

            Assert.Equal("A1:A1", sheet.GetTableRange(query.TableName));
            Assert.Equal("Value", sheet.CellAt(1, 1).GetValue<string>());
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_ModernChart_RemovePreservesSharedChartExRelationship() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            ExcelModernChart chart = sheet.AddModernChart(
                new ExcelChartData(
                    new[] { "A", "B" },
                    new[] { new ExcelChartSeries("Value", new[] { 1d, 2d }) }),
                1,
                1,
                ExcelModernChartType.Funnel,
                "Shared");
            DrawingsPart drawingsPart = sheet.WorksheetPart.DrawingsPart!;
            Xdr.OneCellAnchor original = drawingsPart.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>().Single();
            var duplicate = (Xdr.OneCellAnchor)original.CloneNode(true);
            Xdr.NonVisualDrawingProperties duplicateProperties = duplicate
                .Descendants<Xdr.NonVisualDrawingProperties>().Single();
            duplicateProperties.Id = 99U;
            duplicateProperties.Name = "Shared modern chart copy";
            drawingsPart.WorksheetDrawing.Append(duplicate);
            drawingsPart.WorksheetDrawing.Save();
            ExcelChartDataRange dataRange = chart.DataRange!;

            ExcelModernChart[] wrappers = sheet.ModernCharts.ToArray();
            Assert.Equal(2, wrappers.Length);
            wrappers[0].Remove();

            ExcelModernChart remaining = Assert.Single(sheet.ModernCharts);
            Assert.Equal("Shared", remaining.Title);
            Assert.Single(drawingsPart.ExtendedChartParts);
            Assert.Equal("A", document[dataRange.SheetName]
                .CellAt(dataRange.CategoryStartRow, dataRange.CategoryStartColumn).GetValue<string>());
            Assert.Empty(document.ValidateOpenXml());

            remaining.Remove();
            Assert.Null(sheet.WorksheetPart.DrawingsPart);
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
