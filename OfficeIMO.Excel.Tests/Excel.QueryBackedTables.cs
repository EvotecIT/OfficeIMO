using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public async Task Test_QueryBackedTable_RequiresExplicitHostPolicyAndRefreshesTransactionally() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.QueryBackedTable.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                ExcelQueryBackedTableInfo source = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                    ConnectionName = "SalesQuery",
                    CommandText = "trusted-command",
                    WorksheetName = sheet.Name,
                    StartCell = "B3",
                    TableName = "SalesResults",
                    ColumnNames = new[] { "Region", "Amount" }
                });
                Assert.False(source.IsImported);
                Assert.Equal("B3:C3", source.Range);
                TableDefinitionPart initialTablePart = sheet.WorksheetPart.TableDefinitionParts.Single();
                QueryTableFields initialFields = initialTablePart.QueryTableParts.Single().QueryTable!
                    .QueryTableRefresh!.QueryTableFields!;
                Assert.Equal(new[] { "Region", "Amount" }, initialFields.Elements<QueryTableField>()
                    .Select(field => field.Name!.Value).ToArray());
                Assert.Equal(
                    initialFields.Elements<QueryTableField>().Select(field => field.Id!.Value),
                    initialTablePart.Table!.TableColumns!.Elements<TableColumn>()
                        .Select(column => column.QueryTableFieldId!.Value));
                ExcelFeatureFinding queryFeature = Assert.Single(document.InspectFeatures().FindFeatures("Query-backed tables"));
                Assert.Equal(OfficeFeatureSupportLevel.PartiallyEditable, queryFeature.SupportLevel);
                Assert.Empty(document.InspectFeatures().FindFeatures("Connections and query tables"));
                var authoredErrors = document.ValidateOpenXml().ToArray();
                Assert.True(authoredErrors.Length == 0, string.Join(Environment.NewLine, authoredErrors));

                var host = new StubQueryHost(new ExcelQueryExecutionResult(
                    new[] { "Territory", "Amount", "AsOf" },
                    new IReadOnlyList<object?>[] {
                        new object?[] { "East", 10d, new DateTime(2026, 1, 2) },
                        new object?[] { "West", 20d, new DateTime(2026, 1, 3) }
                    }));
                await Assert.ThrowsAsync<SecurityException>(() => document.RefreshQueryAsync(
                    source.ConnectionName,
                    host,
                    new ExcelQueryExecutionPolicy()));
                Assert.Equal(0, host.CallCount);

                ExcelQueryRefreshResult refreshed = await document.RefreshQueryAsync(
                    source.ConnectionName,
                    host,
                    new ExcelQueryExecutionPolicy { AllowExecution = true });
                Assert.Equal(2, refreshed.RowCount);
                Assert.Equal(3, refreshed.ColumnCount);
                Assert.Equal("B3:D5", refreshed.Range);
                Assert.Equal("Territory", sheet.CellAt(3, 2).GetValue<string>());
                Assert.Equal("East", sheet.CellAt(4, 2).GetValue<string>());
                Assert.Equal(20d, sheet.CellAt(5, 3).GetValue<double>());
                QueryTableFields refreshedFields = initialTablePart.QueryTableParts.Single().QueryTable!
                    .QueryTableRefresh!.QueryTableFields!;
                Assert.Equal(new[] { "Territory", "Amount", "AsOf" }, refreshedFields.Elements<QueryTableField>()
                    .Select(field => field.Name!.Value).ToArray());
                Assert.Equal(
                    refreshedFields.Elements<QueryTableField>().Select(field => field.Id!.Value),
                    initialTablePart.Table!.TableColumns!.Elements<TableColumn>()
                        .Select(column => column.QueryTableFieldId!.Value));
                Assert.Empty(document.ValidateOpenXml());

                var oversized = new StubQueryHost(new ExcelQueryExecutionResult(
                    new[] { "Only" },
                    new IReadOnlyList<object?>[] { new object?[] { "one" }, new object?[] { "two" } }));
                await Assert.ThrowsAsync<InvalidOperationException>(() => document.RefreshQueryAsync(
                    source.TableName,
                    oversized,
                    new ExcelQueryExecutionPolicy { AllowExecution = true, MaximumRows = 1 }));
                Assert.Equal("B3:D5", sheet.GetTableRange(source.TableName));
                Assert.Equal("East", sheet.CellAt(4, 2).GetValue<string>());
                document.Save();
            }

            using (ExcelDocument imported = ExcelDocument.Load(filePath)) {
                ExcelQueryBackedTableInfo source = Assert.Single(imported.GetQueryBackedTables());
                Assert.True(source.IsImported);
                var host = new StubQueryHost(new ExcelQueryExecutionResult(
                    new[] { "Region" },
                    new IReadOnlyList<object?>[] { new object?[] { "North" } }));
                await Assert.ThrowsAsync<SecurityException>(() => imported.RefreshQueryAsync(
                    source.ConnectionName,
                    host,
                    new ExcelQueryExecutionPolicy { AllowExecution = true }));
                Assert.Equal(0, host.CallCount);

                ExcelQueryRefreshResult refreshed = await imported.RefreshQueryAsync(
                    source.ConnectionName,
                    host,
                    new ExcelQueryExecutionPolicy { AllowExecution = true, AllowImportedCommands = true });
                Assert.Equal("B3:B4", refreshed.Range);
                Assert.Equal("North", imported["Data"].CellAt(4, 2).GetValue<string>());
                var importedErrors = imported.ValidateOpenXml().ToArray();
                Assert.True(importedErrors.Length == 0, string.Join(Environment.NewLine, importedErrors));
                Assert.True(imported.RemoveQueryBackedTable(source.ConnectionName, preserveTable: false));
                Assert.Empty(imported.GetQueryBackedTables());
                Assert.Null(imported["Data"].GetTableRange(source.TableName));
                Assert.Equal("North", imported["Data"].CellAt(4, 2).GetValue<string>());
                var removedErrors = imported.ValidateOpenXml().ToArray();
                Assert.True(removedErrors.Length == 0, string.Join(Environment.NewLine, removedErrors));
            }
        }

        [Fact]
        public async Task Test_QueryBackedTable_CancellationStopsLazyMaterializationBeforeMutation() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            ExcelQueryBackedTableInfo source = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "LazyQuery",
                WorksheetName = sheet.Name,
                TableName = "LazyResults",
                ColumnNames = new[] { "Value" }
            });
            using var cancellation = new CancellationTokenSource();
            IEnumerable<IReadOnlyList<object?>> Rows() {
                yield return new object?[] { "first" };
                cancellation.Cancel();
                yield return new object?[] { "second" };
            }
            var host = new StubQueryHost(new ExcelQueryExecutionResult(new[] { "Value" }, Rows()));

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => document.RefreshQueryAsync(
                source.ConnectionName,
                host,
                new ExcelQueryExecutionPolicy { AllowExecution = true },
                cancellation.Token));
            Assert.Equal("A1:A1", sheet.GetTableRange(source.TableName));
            Assert.Equal("Value", sheet.CellAt(1, 1).GetValue<string>());

            var unsupported = new StubQueryHost(new ExcelQueryExecutionResult(
                new[] { "Value" },
                new IReadOnlyList<object?>[] { new object?[] { new object() } }));
            await Assert.ThrowsAsync<InvalidDataException>(() => document.RefreshQueryAsync(
                source.ConnectionName,
                unsupported,
                new ExcelQueryExecutionPolicy { AllowExecution = true }));
            Assert.Equal("A1:A1", sheet.GetTableRange(source.TableName));

            var headerOnly = new StubQueryHost(new ExcelQueryExecutionResult(
                new[] { "LongHeader" },
                Array.Empty<IReadOnlyList<object?>>()));
            await Assert.ThrowsAsync<InvalidOperationException>(() => document.RefreshQueryAsync(
                source.ConnectionName,
                headerOnly,
                new ExcelQueryExecutionPolicy { AllowExecution = true, MaximumCharacters = 5 }));
            Assert.Equal("A1:A1", sheet.GetTableRange(source.TableName));
        }

        [Fact]
        public async Task Test_QueryBackedTable_SnapshotsPolicyAndParticipatesInStructuralEdits() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            ExcelQueryBackedTableInfo source = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "SafeQuery",
                WorksheetName = sheet.Name,
                TableName = "SafeResults",
                ColumnNames = new[] { "Name", "Value" }
            });
            QueryTablePart queryPart = sheet.WorksheetPart.TableDefinitionParts.Single().QueryTableParts.Single();
            queryPart.QueryTable!.QueryTableRefresh!.Append(new SortState { Reference = "A1:B1" });
            queryPart.QueryTable.Save();

            var host = new MutatingPolicyHost(new ExcelQueryExecutionResult(
                new[] { "Name", "Value" },
                new IReadOnlyList<object?>[] { new object?[] { "kept", 1d } }));
            ExcelQueryRefreshResult refreshed = await document.RefreshQueryAsync(
                source.ConnectionName,
                host,
                new ExcelQueryExecutionPolicy { AllowExecution = true, MaximumRows = 1 });
            Assert.Equal(1, refreshed.RowCount);

            sheet.InsertRows(1, 2);
            Assert.Equal("A3:B4", sheet.GetTableRange(source.TableName));
            Assert.Equal("A3:B3", queryPart.QueryTable.QueryTableRefresh.SortState!.Reference!.Value);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public async Task Test_QueryBackedTable_RejectsOccupiedExpansionAndRefreshesSharedConnectionsByTable() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet firstSheet = document.AddWorksheet("First");
            ExcelSheet secondSheet = document.AddWorksheet("Second");
            ExcelQueryBackedTableInfo first = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "SharedQuery",
                WorksheetName = firstSheet.Name,
                TableName = "FirstResults",
                ColumnNames = new[] { "Value" }
            });
            ExcelQueryBackedTableInfo second = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "TemporaryQuery",
                WorksheetName = secondSheet.Name,
                TableName = "SecondResults",
                ColumnNames = new[] { "Value" }
            });

            firstSheet.CellValue(2, 2, "protected");
            var growing = new StubQueryHost(new ExcelQueryExecutionResult(
                new[] { "Value", "Extra" },
                new IReadOnlyList<object?>[] { new object?[] { "new", "collision" } }));
            await Assert.ThrowsAsync<InvalidOperationException>(() => document.RefreshQueryAsync(
                first.TableName,
                growing,
                new ExcelQueryExecutionPolicy { AllowExecution = true }));
            Assert.Equal("protected", firstSheet.CellAt(2, 2).GetValue<string>());
            Assert.Equal("A1:A1", firstSheet.GetTableRange(first.TableName));

            QueryTablePart firstQueryPart = firstSheet.WorksheetPart.TableDefinitionParts.Single().QueryTableParts.Single();
            QueryTablePart secondQueryPart = secondSheet.WorksheetPart.TableDefinitionParts.Single().QueryTableParts.Single();
            secondQueryPart.QueryTable!.ConnectionId = first.ConnectionId;
            secondQueryPart.QueryTable.Save();
            ConnectionsPart connectionsPart = document._spreadSheetDocument.WorkbookPart!.ConnectionsPart!;
            connectionsPart.Connections!.Elements<Connection>()
                .Single(connection => connection.Id!.Value == second.ConnectionId).Remove();
            connectionsPart.Connections.Save();

            var host = new RecordingQueryHost();
            IReadOnlyList<ExcelQueryRefreshResult> refreshed = await document.RefreshQueriesAsync(
                host,
                new ExcelQueryExecutionPolicy { AllowExecution = true });
            Assert.Equal(2, refreshed.Count);
            Assert.Equal(new[] { "FirstResults", "SecondResults" }, host.TableNames.OrderBy(name => name).ToArray());
            Assert.Equal("FirstResults", firstSheet.CellAt(2, 1).GetValue<string>());
            Assert.Equal("SecondResults", secondSheet.CellAt(2, 1).GetValue<string>());
            await Assert.ThrowsAsync<InvalidOperationException>(() => document.RefreshQueryAsync(
                first.ConnectionName,
                host,
                new ExcelQueryExecutionPolicy { AllowExecution = true }));

            QueryTablePart worksheetConsumer = firstSheet.WorksheetPart.AddNewPart<QueryTablePart>();
            worksheetConsumer.QueryTable = (QueryTable)firstQueryPart.QueryTable!.CloneNode(true);
            worksheetConsumer.QueryTable.Save();
            ExcelSheet mappedSheet = document.AddWorksheet("Mapped");
            mappedSheet.CellValue(1, 1, "Value");
            mappedSheet.CellValue(2, 1, "kept");
            mappedSheet.AddTable("A1:A2", true, "MappedResults", OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium2);
            TableDefinitionPart mappedTablePart = mappedSheet.WorksheetPart.TableDefinitionParts.Single();
            mappedTablePart.Table!.ConnectionId = first.ConnectionId;
            mappedTablePart.Table.Save();
            Assert.True(document.RemoveQueryBackedTable(first.TableName));
            Assert.True(document.RemoveQueryBackedTable(second.TableName));
            Assert.Contains(connectionsPart.Connections!.Elements<Connection>(), connection =>
                connection.Id?.Value == first.ConnectionId);
            Assert.Equal(first.ConnectionId, mappedTablePart.Table.ConnectionId!.Value);
            Assert.Empty(document.ValidateOpenXml());
        }

        private sealed class StubQueryHost : IExcelQueryExecutionHost {
            private readonly ExcelQueryExecutionResult _result;

            internal StubQueryHost(ExcelQueryExecutionResult result) {
                _result = result;
            }

            internal int CallCount { get; private set; }

            public Task<ExcelQueryExecutionResult> ExecuteAsync(
                ExcelQueryExecutionRequest request,
                CancellationToken cancellationToken) {
                CallCount++;
                return Task.FromResult(_result);
            }
        }

        private sealed class MutatingPolicyHost : IExcelQueryExecutionHost {
            private readonly ExcelQueryExecutionResult _result;

            internal MutatingPolicyHost(ExcelQueryExecutionResult result) {
                _result = result;
            }

            public Task<ExcelQueryExecutionResult> ExecuteAsync(
                ExcelQueryExecutionRequest request,
                CancellationToken cancellationToken) {
                request.Policy.MaximumRows = 0;
                request.Policy.MaximumCells = 0;
                return Task.FromResult(_result);
            }
        }

        private sealed class RecordingQueryHost : IExcelQueryExecutionHost {
            internal List<string> TableNames { get; } = new List<string>();

            public Task<ExcelQueryExecutionResult> ExecuteAsync(
                ExcelQueryExecutionRequest request,
                CancellationToken cancellationToken) {
                TableNames.Add(request.TableName);
                return Task.FromResult(new ExcelQueryExecutionResult(
                    new[] { "Value" },
                    new IReadOnlyList<object?>[] { new object?[] { request.TableName } }));
            }
        }
    }
}
