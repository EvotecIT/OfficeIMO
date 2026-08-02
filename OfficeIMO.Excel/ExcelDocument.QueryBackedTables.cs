using System.Globalization;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private readonly HashSet<uint> _authoredQueryConnectionIds = new HashSet<uint>();

        /// <summary>Creates a native query-backed worksheet table without executing its command.</summary>
        public ExcelQueryBackedTableInfo AddQueryBackedTable(ExcelQueryBackedTableOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (string.IsNullOrWhiteSpace(options.ConnectionName)) throw new ArgumentNullException(nameof(options.ConnectionName));
            if (string.IsNullOrWhiteSpace(options.WorksheetName)) throw new ArgumentNullException(nameof(options.WorksheetName));
            if (string.IsNullOrWhiteSpace(options.TableName)) throw new ArgumentNullException(nameof(options.TableName));
            string[] columns = ValidateQueryColumns(options.ColumnNames);
            (int row, int column) = A1.ParseCellRef(options.StartCell);
            if (row < 1 || column < 1 || column + columns.Length - 1 > 16_384) {
                throw new ArgumentOutOfRangeException(nameof(options.StartCell));
            }
            string connectionName = options.ConnectionName.Trim();
            string requestedTableName = options.TableName.Trim();

            ExcelSheet sheet = this[options.WorksheetName];
            string range = A1.CellReference(row, column) + ":" + A1.CellReference(row, column + columns.Length - 1);
            string? tableName = null;
            uint connectionId = 0U;
            sheet.ApplyTransactionalMutation(_ => {
                if (HasQueryBackedTableNameConflict(connectionName, requestedTableName)) {
                    throw new InvalidOperationException("Query connection and table names must be unique.");
                }
                for (int index = 0; index < columns.Length; index++) {
                    sheet.CellValue(row, column + index, columns[index]);
                }
                tableName = sheet.AddTableAndGetName(
                    range,
                    hasHeader: true,
                    requestedTableName,
                    TableStyle.TableStyleMedium2,
                    includeAutoFilter: true,
                    validationMode: TableNameValidationMode.Strict,
                    headerNames: columns);

                connectionId = GetNextPowerQueryConnectionId();
                Connection connection = CreateNativeQueryConnection(
                    connectionId,
                    connectionName,
                    options.Description,
                    options.CommandText,
                    options.RefreshOnOpen);
                AppendNativeQueryConnection(connection);

                TableDefinitionPart tablePart = sheet.WorksheetPart.TableDefinitionParts.Single(part =>
                    string.Equals(part.Table?.Name?.Value ?? part.Table?.DisplayName?.Value, tableName, StringComparison.OrdinalIgnoreCase));
                QueryTablePart queryPart = tablePart.AddNewPart<QueryTablePart>();
                queryPart.QueryTable = CreateNativeQueryTable(
                    tableName!,
                    connectionId,
                    tablePart.Table!.TableColumns!,
                    columns);
                queryPart.QueryTable.Save();
                return columns.Length;
            }, new ExcelMutationPlanOptions(), CancellationToken.None);
            Locking.ExecuteWrite(EnsureLock(), () => {
                _authoredQueryConnectionIds.Add(connectionId);
                MarkMetadataPartChanged();
            });

            return new ExcelQueryBackedTableInfo(
                connectionId,
                connectionName,
                options.CommandText ?? string.Empty,
                sheet.Name,
                tableName!,
                range,
                imported: false);
        }

        private bool HasQueryBackedTableNameConflict(string connectionName, string tableName) {
            IReadOnlyDictionary<uint, (string Name, string Command)> connections = ReadNativeQueryConnections();
            foreach (WorksheetPart worksheetPart in WorkbookPartRoot.WorksheetParts) {
                foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) {
                    Table? table = tablePart.Table;
                    QueryTablePart? queryPart = tablePart.QueryTableParts.FirstOrDefault();
                    if (table == null
                        || queryPart?.QueryTable?.ConnectionId?.Value is not uint connectionId
                        || !connections.TryGetValue(connectionId, out var connection)) {
                        continue;
                    }

                    string existingTableName = table.Name?.Value ?? table.DisplayName?.Value ?? string.Empty;
                    if (string.Equals(connection.Name, connectionName, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(existingTableName, tableName, StringComparison.OrdinalIgnoreCase)) {
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>Lists query-backed worksheet tables that have a resolvable native query-table relationship.</summary>
        public IReadOnlyList<ExcelQueryBackedTableInfo> GetQueryBackedTables() {
            IReadOnlyDictionary<uint, (string Name, string Command)> connections = ReadNativeQueryConnections();
            var results = new List<ExcelQueryBackedTableInfo>();
            foreach (ExcelSheet sheet in Sheets) {
                foreach (TableDefinitionPart tablePart in sheet.WorksheetPart.TableDefinitionParts) {
                    Table? table = tablePart.Table;
                    if (table == null) continue;
                    QueryTablePart? queryPart = tablePart.QueryTableParts.FirstOrDefault();
                    if (queryPart?.QueryTable?.ConnectionId?.Value is not uint connectionId) continue;
                    if (!connections.TryGetValue(connectionId, out var connection)) continue;
                    string tableName = table.Name?.Value ?? table.DisplayName?.Value ?? string.Empty;
                    results.Add(new ExcelQueryBackedTableInfo(
                        connectionId,
                        connection.Name ?? string.Empty,
                        connection.Command ?? string.Empty,
                        sheet.Name,
                        tableName,
                        table.Reference?.Value ?? string.Empty,
                        imported: !_authoredQueryConnectionIds.Contains(connectionId)));
                }
            }
            return results.OrderBy(item => item.WorksheetName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.TableName, StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        /// <summary>Executes one query through an explicit caller host and atomically replaces its owned table data.</summary>
        public async Task<ExcelQueryRefreshResult> RefreshQueryAsync(
            string connectionOrTableName,
            IExcelQueryExecutionHost host,
            ExcelQueryExecutionPolicy policy,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(connectionOrTableName)) throw new ArgumentNullException(nameof(connectionOrTableName));
            if (host == null) throw new ArgumentNullException(nameof(host));
            ExcelQueryExecutionPolicy effectivePolicy = SnapshotQueryExecutionPolicy(policy);
            ValidateQueryExecutionPolicy(effectivePolicy);
            ExcelQueryBackedTableInfo[] matches = GetQueryBackedTables().Where(item =>
                string.Equals(item.ConnectionName, connectionOrTableName.Trim(), StringComparison.OrdinalIgnoreCase)
                || string.Equals(item.TableName, connectionOrTableName.Trim(), StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matches.Length == 0) throw new InvalidOperationException($"Query-backed table '{connectionOrTableName}' was not found.");
            if (matches.Length > 1) throw new InvalidOperationException(
                $"Query-backed table or connection '{connectionOrTableName}' is ambiguous; use a unique table name.");
            return await RefreshQuerySourceAsync(matches[0], host, effectivePolicy, cancellationToken).ConfigureAwait(false);
        }

        private async Task<ExcelQueryRefreshResult> RefreshQuerySourceAsync(
            ExcelQueryBackedTableInfo source,
            IExcelQueryExecutionHost host,
            ExcelQueryExecutionPolicy effectivePolicy,
            CancellationToken cancellationToken) {
            if (source.IsImported && !effectivePolicy.AllowImportedCommands) {
                throw new SecurityException("Execution of commands loaded from an imported workbook is disabled by policy.");
            }

            cancellationToken.ThrowIfCancellationRequested();
            var request = new ExcelQueryExecutionRequest(
                source.ConnectionId,
                source.ConnectionName,
                source.CommandText,
                source.WorksheetName,
                source.TableName,
                source.IsImported,
                SnapshotQueryExecutionPolicy(effectivePolicy));
            ExcelQueryExecutionResult result = await host.ExecuteAsync(request, cancellationToken).ConfigureAwait(false)
                ?? throw new InvalidDataException("The query host returned no result.");
            MaterializedQueryResult materialized = MaterializeQueryResult(result, effectivePolicy, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();

            ExcelSheet sheet = this[source.WorksheetName];
            string updatedRange = sheet.ReplaceQueryBackedTableData(
                source.TableName,
                materialized.Columns,
                materialized.Rows,
                cancellationToken);
            MarkMetadataPartChanged();
            ExcelQueryBackedTableInfo refreshed = GetQueryBackedTables().Single(item =>
                item.ConnectionId == source.ConnectionId
                && string.Equals(item.WorksheetName, source.WorksheetName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(item.TableName, source.TableName, StringComparison.OrdinalIgnoreCase));
            return new ExcelQueryRefreshResult(refreshed, materialized.Rows.Count, materialized.Columns.Length, updatedRange);
        }

        /// <summary>Refreshes all query-backed tables sequentially through one explicit caller host and policy.</summary>
        public async Task<IReadOnlyList<ExcelQueryRefreshResult>> RefreshQueriesAsync(
            IExcelQueryExecutionHost host,
            ExcelQueryExecutionPolicy policy,
            CancellationToken cancellationToken = default) {
            if (host == null) throw new ArgumentNullException(nameof(host));
            ExcelQueryExecutionPolicy effectivePolicy = SnapshotQueryExecutionPolicy(policy);
            ValidateQueryExecutionPolicy(effectivePolicy);
            ExcelQueryBackedTableInfo[] sources = GetQueryBackedTables().ToArray();
            var results = new List<ExcelQueryRefreshResult>(sources.Length);
            foreach (ExcelQueryBackedTableInfo source in sources) {
                cancellationToken.ThrowIfCancellationRequested();
                results.Add(await RefreshQuerySourceAsync(
                    source,
                    host,
                    SnapshotQueryExecutionPolicy(effectivePolicy),
                    cancellationToken).ConfigureAwait(false));
            }
            return results;
        }

        /// <summary>Detaches a query binding and removes its unused OfficeIMO-owned connection, optionally converting the table to a normal cell range.</summary>
        public bool RemoveQueryBackedTable(string connectionOrTableName, bool preserveTable = true) {
            if (string.IsNullOrWhiteSpace(connectionOrTableName)) throw new ArgumentNullException(nameof(connectionOrTableName));
            ExcelQueryBackedTableInfo[] matches = GetQueryBackedTables().Where(item =>
                string.Equals(item.ConnectionName, connectionOrTableName.Trim(), StringComparison.OrdinalIgnoreCase)
                || string.Equals(item.TableName, connectionOrTableName.Trim(), StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matches.Length == 0) return false;
            if (matches.Length > 1) throw new InvalidOperationException(
                $"Query-backed table or connection '{connectionOrTableName}' is ambiguous; use a unique table name.");
            ExcelQueryBackedTableInfo source = matches[0];
            ExcelSheet sheet = this[source.WorksheetName];
            sheet.RemoveQueryBackedTableBinding(source.TableName, preserveTable);

            RemoveUnusedAuthoredQueryConnections(new[] { source.ConnectionId });
            MarkMetadataPartChanged();
            return true;
        }

        internal IReadOnlyList<uint> GetWorksheetQueryConnectionIds(WorksheetPart worksheetPart) {
            var result = new HashSet<uint>();
            foreach (QueryTablePart queryPart in ExcelPackageQueryTableParts.Enumerate(worksheetPart)) {
                if (queryPart.QueryTable?.ConnectionId?.Value is uint connectionId) result.Add(connectionId);
            }
            foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) {
                if (tablePart.Table?.ConnectionId?.Value is uint connectionId) result.Add(connectionId);
            }
            foreach (SingleXmlCell cell in worksheetPart.SingleCellTablePart?.SingleXmlCells?
                .Elements<SingleXmlCell>() ?? Enumerable.Empty<SingleXmlCell>()) {
                if (cell.ConnectionId?.Value is uint connectionId) result.Add(connectionId);
            }
            return result.ToArray();
        }

        internal void RemoveUnusedAuthoredQueryConnections(IEnumerable<uint> connectionIds) {
            bool changed = false;
            foreach (uint connectionId in connectionIds.Distinct()) {
                if (!_authoredQueryConnectionIds.Contains(connectionId)
                    || IsNativeConnectionReferenced(connectionId)) continue;
                RemoveNativeQueryConnection(connectionId);
                _authoredQueryConnectionIds.Remove(connectionId);
                changed = true;
            }
            if (!changed) return;
            MarkMetadataPartChanged();
        }

        private void AppendNativeQueryConnection(Connection connection) {
            OpenXmlPart? part = GetWorkbookConnectionPart();
            if (part == null) {
                ConnectionsPart connectionsPart = WorkbookPartRoot.AddNewPart<ConnectionsPart>();
                connectionsPart.Connections = new Connections(connection);
                connectionsPart.Connections.Save();
                return;
            }

            if (part is ConnectionsPart nativePart) {
                nativePart.Connections ??= new Connections();
                nativePart.Connections.Append(connection);
                nativePart.Connections.Save();
                return;
            }

            string merged = MergeWorkbookConnectionMetadata(ReadOpenXmlPartText(part), connection.OuterXml);
            WriteOpenXmlPartText(part, merged);
        }

        private void RemoveNativeQueryConnection(uint connectionId) {
            foreach (OpenXmlPart part in EnumerateWorkbookConnectionParts().ToArray()) {
                if (part is ConnectionsPart nativePart) {
                    Connection? connection = nativePart.Connections?.Elements<Connection>()
                        .FirstOrDefault(item => item.Id?.Value == connectionId);
                    if (connection == null) continue;
                    connection.Remove();
                    if (nativePart.Connections?.Elements<Connection>().Any() == true) nativePart.Connections.Save();
                    else WorkbookPartRoot.DeletePart(nativePart);
                    continue;
                }

                try {
                    XDocument document = XDocument.Parse(ReadOpenXmlPartText(part), LoadOptions.PreserveWhitespace);
                    XElement[] matches = document.Descendants()
                        .Where(element => element.Name.LocalName == "connection"
                            && uint.TryParse(element.Attribute("id")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint id)
                            && id == connectionId)
                        .ToArray();
                    if (matches.Length == 0) continue;
                    foreach (XElement match in matches) match.Remove();
                    if (document.Root != null) {
                        document.Root.SetAttributeValue(
                            "count",
                            document.Root.Elements().Count(element => element.Name.LocalName == "connection")
                                .ToString(CultureInfo.InvariantCulture));
                    }
                    WriteOpenXmlPartText(part, document.ToString(SaveOptions.DisableFormatting));
                } catch (InvalidDataException) {
                    // Preserve oversized or malformed caller-owned metadata instead of mutating it.
                } catch (IOException) {
                    // Preserve unreadable caller-owned metadata instead of mutating it.
                } catch (System.Xml.XmlException) {
                    // Preserve malformed caller-owned metadata instead of mutating it.
                }
            }
        }

        private bool IsNativeConnectionReferenced(uint connectionId) {
            if (WorkbookPartRoot.WorksheetParts
                .SelectMany(ExcelPackageQueryTableParts.Enumerate)
                .Any(part => part.QueryTable?.ConnectionId?.Value == connectionId)) return true;
            if (WorkbookPartRoot.WorksheetParts
                .SelectMany(part => part.TableDefinitionParts)
                .Any(part => part.Table?.ConnectionId?.Value == connectionId)) return true;
            if (WorkbookPartRoot.PivotTableCacheDefinitionParts.Any(part =>
                part.PivotCacheDefinition?.CacheSource?.ConnectionId?.Value == connectionId)) return true;
            if (WorkbookPartRoot.WorksheetParts.Any(part => part.SingleCellTablePart?.SingleXmlCells?
                .Elements<SingleXmlCell>().Any(cell => cell.ConnectionId?.Value == connectionId) == true)) return true;
            return WorkbookPartRoot.CustomXmlMappingsPart?.MapInfo?
                .Descendants<DataBinding>().Any(binding => binding.ConnectionId?.Value == connectionId) == true;
        }

        private static void ValidateQueryExecutionPolicy(ExcelQueryExecutionPolicy policy) {
            if (policy == null) throw new ArgumentNullException(nameof(policy));
            if (!policy.AllowExecution) throw new SecurityException("Query execution must be enabled explicitly by policy.");
            if (policy.MaximumRows < 0 || policy.MaximumColumns < 1 || policy.MaximumCells < 0 || policy.MaximumCharacters < 0) {
                throw new ArgumentOutOfRangeException(nameof(policy), "Query execution budgets must be non-negative and allow at least one column.");
            }
        }

        private static ExcelQueryExecutionPolicy SnapshotQueryExecutionPolicy(ExcelQueryExecutionPolicy policy) {
            if (policy == null) throw new ArgumentNullException(nameof(policy));
            return new ExcelQueryExecutionPolicy {
                AllowExecution = policy.AllowExecution,
                AllowImportedCommands = policy.AllowImportedCommands,
                MaximumRows = policy.MaximumRows,
                MaximumColumns = policy.MaximumColumns,
                MaximumCells = policy.MaximumCells,
                MaximumCharacters = policy.MaximumCharacters
            };
        }

        private MaterializedQueryResult MaterializeQueryResult(
            ExcelQueryExecutionResult result,
            ExcelQueryExecutionPolicy policy,
            CancellationToken cancellationToken) {
            string[] columns = ValidateQueryColumns(result.ColumnNames);
            if (columns.Length > policy.MaximumColumns) {
                throw new InvalidOperationException($"Query result exceeds MaximumColumns ({policy.MaximumColumns}).");
            }
            long characters = columns.Sum(column => (long)column.Length);
            if (characters > policy.MaximumCharacters) {
                throw new InvalidOperationException($"Query result exceeds MaximumCharacters ({policy.MaximumCharacters}).");
            }
            var rows = new List<object?[]>();
            foreach (IReadOnlyList<object?> row in result.Rows) {
                cancellationToken.ThrowIfCancellationRequested();
                if (rows.Count >= policy.MaximumRows) {
                    throw new InvalidOperationException($"Query result exceeds MaximumRows ({policy.MaximumRows}).");
                }
                if (row == null || row.Count != columns.Length) {
                    throw new InvalidDataException("Every query result row must match the declared column count.");
                }
                long nextCellCount = checked((long)(rows.Count + 1) * columns.Length);
                if (nextCellCount > policy.MaximumCells) {
                    throw new InvalidOperationException($"Query result exceeds MaximumCells ({policy.MaximumCells}).");
                }
                var values = new object?[columns.Length];
                for (int index = 0; index < columns.Length; index++) {
                    object? value = NormalizeQueryCellValue(row[index]);
                    if (value is string text) {
                        if (text.Length > 32_767) {
                            throw new InvalidDataException("Query result cell text exceeds Excel's 32,767-character limit.");
                        }
                        characters = checked(characters + text.Length);
                        if (characters > policy.MaximumCharacters) {
                            throw new InvalidOperationException($"Query result exceeds MaximumCharacters ({policy.MaximumCharacters}).");
                        }
                    }
                    values[index] = value;
                }
                rows.Add(values);
            }
            return new MaterializedQueryResult(columns, rows);
        }

        private object? NormalizeQueryCellValue(object? value) {
            if (value == null || value == DBNull.Value) return null;
            switch (value) {
                case double number when double.IsNaN(number) || double.IsInfinity(number):
                    throw new InvalidDataException("Query result contains a non-finite floating-point value.");
                case float number when float.IsNaN(number) || float.IsInfinity(number):
                    throw new InvalidDataException("Query result contains a non-finite floating-point value.");
                case string:
                case double:
                case float:
                case decimal:
                case int:
                case long:
                case short:
                case uint:
                case ulong:
                case ushort:
                case byte:
                case sbyte:
                case bool:
                case DateTime:
                case TimeSpan:
                    return value;
                case DateTimeOffset dateTimeOffset:
                    return CoerceValueHelper.TryConvertDateTimeOffset(
                        dateTimeOffset,
                        DateTimeOffsetWriteStrategy,
                        DateSystem,
                        out DateTime converted,
                        out _)
                            ? converted
                            : dateTimeOffset.ToString("o", CultureInfo.InvariantCulture);
#if NET6_0_OR_GREATER
                case DateOnly:
                case TimeOnly:
                    return value;
#endif
                case Guid guid:
                    return guid.ToString("D");
                case char character:
                    return character.ToString();
                default:
                    throw new InvalidDataException(
                        $"Query result value type '{value.GetType().FullName}' is not a supported Excel cell scalar.");
            }
        }

        private static string[] ValidateQueryColumns(IReadOnlyList<string> columnNames) {
            if (columnNames == null || columnNames.Count == 0) {
                throw new ArgumentException("A query-backed table requires at least one column.", nameof(columnNames));
            }
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var columns = new string[columnNames.Count];
            for (int index = 0; index < columnNames.Count; index++) {
                string name = columnNames[index]?.Trim() ?? string.Empty;
                if (name.Length == 0) throw new ArgumentException($"Query column {index + 1} has no name.", nameof(columnNames));
                if (name.Length > 32_767) throw new ArgumentException($"Query column {index + 1} exceeds Excel's cell-text limit.", nameof(columnNames));
                try {
                    XmlConvert.VerifyXmlChars(name);
                } catch (XmlException exception) {
                    throw new ArgumentException($"Query column {index + 1} contains characters that are invalid in XML.", nameof(columnNames), exception);
                }
                if (!used.Add(name)) throw new ArgumentException($"Query column '{name}' is duplicated.", nameof(columnNames));
                columns[index] = name;
            }
            return columns;
        }

        private IReadOnlyDictionary<uint, (string Name, string Command)> ReadNativeQueryConnections() {
            var result = new Dictionary<uint, (string Name, string Command)>();
            foreach (OpenXmlPart part in EnumerateWorkbookConnectionParts()) {
                try {
                    XDocument xml = XDocument.Parse(ReadOpenXmlPartText(part));
                    foreach (XElement connection in xml.Descendants().Where(element => element.Name.LocalName == "connection")) {
                        if (!uint.TryParse(connection.Attribute("id")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint id)) continue;
                        string name = connection.Attribute("name")?.Value ?? string.Empty;
                        XElement? database = connection.Elements().FirstOrDefault(element => element.Name.LocalName == "dbPr");
                        result[id] = (name, database?.Attribute("command")?.Value ?? string.Empty);
                    }
                } catch {
                    // Caller-supplied extended metadata may not be a parseable connection document.
                }
            }
            return result;
        }

        private static Connection CreateNativeQueryConnection(
            uint connectionId,
            string name,
            string? description,
            string? commandText,
            bool refreshOnOpen) {
            var connection = new Connection {
                Id = connectionId,
                Name = name,
                Type = 5U,
                RefreshedVersion = 7,
                RefreshOnLoad = refreshOnOpen
            };
            if (!string.IsNullOrWhiteSpace(description)) connection.Description = description!.Trim();
            if (!string.IsNullOrWhiteSpace(commandText)) {
                connection.DatabaseProperties = new DatabaseProperties {
                    Connection = "Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" + name + ";Extended Properties=\"\"",
                    Command = commandText,
                    CommandType = 1U
                };
            }
            return connection;
        }

        private static QueryTable CreateNativeQueryTable(
            string tableName,
            uint connectionId,
            TableColumns tableColumns,
            IReadOnlyList<string> columnNames) {
            var refresh = new QueryTableRefresh(new QueryTableFields()) {
                MinimumVersion = 0,
                NextId = 1U
            };
            var queryTable = new QueryTable(refresh) {
                Name = tableName,
                ConnectionId = connectionId,
                AutoFormatId = 16U,
                ApplyNumberFormats = false,
                ApplyBorderFormats = false,
                ApplyFontFormats = false,
                ApplyPatternFormats = false,
                ApplyAlignmentFormats = false,
                ApplyWidthHeightFormats = false
            };
            SynchronizeNativeQueryFields(queryTable, tableColumns, columnNames);
            return queryTable;
        }

        internal static void SynchronizeNativeQueryFields(
            QueryTable queryTable,
            TableColumns tableColumns,
            IReadOnlyList<string> columnNames) {
            QueryTableRefresh refresh = queryTable.QueryTableRefresh
                ?? queryTable.AppendChild(new QueryTableRefresh { MinimumVersion = 0, NextId = 1U });
            QueryTableFields fields = refresh.QueryTableFields
                ?? refresh.PrependChild(new QueryTableFields());
            List<QueryTableField> existing = fields.Elements<QueryTableField>().ToList();
            List<TableColumn> columns = tableColumns.Elements<TableColumn>().ToList();
            uint nextId = existing.Select(field => field.Id?.Value ?? 0U).DefaultIfEmpty(0U).Max() + 1U;
            var usedIds = new HashSet<uint>();
            for (int index = 0; index < columnNames.Count; index++) {
                QueryTableField field = index < existing.Count
                    ? existing[index]
                    : fields.AppendChild(new QueryTableField());
                uint fieldId = field.Id?.Value is uint candidate && candidate > 0U && usedIds.Add(candidate)
                    ? candidate
                    : nextId++;
                usedIds.Add(fieldId);
                field.Id = fieldId;
                field.Name = columnNames[index];
                field.TableColumnId = columns[index].Id?.Value ?? (uint)(index + 1);
                columns[index].QueryTableFieldId = fieldId;
            }
            for (int index = existing.Count - 1; index >= columnNames.Count; index--) existing[index].Remove();
            fields.Count = (uint)columnNames.Count;
            refresh.NextId = usedIds.Count == 0 ? 1U : usedIds.Max() + 1U;
        }

        private sealed class MaterializedQueryResult {
            internal MaterializedQueryResult(string[] columns, List<object?[]> rows) {
                Columns = columns;
                Rows = rows;
            }

            internal string[] Columns { get; }
            internal List<object?[]> Rows { get; }
        }
    }
}
