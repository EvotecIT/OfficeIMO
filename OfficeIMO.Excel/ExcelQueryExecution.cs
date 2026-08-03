using System.Collections.ObjectModel;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    /// <summary>Explicit security and materialization policy for caller-hosted query execution.</summary>
    public sealed class ExcelQueryExecutionPolicy {
        /// <summary>Must be enabled explicitly before OfficeIMO invokes a query host.</summary>
        public bool AllowExecution { get; set; }

        /// <summary>Allows commands read from an imported workbook to be sent to the host.</summary>
        public bool AllowImportedCommands { get; set; }

        /// <summary>Maximum result rows, excluding the header row.</summary>
        public int MaximumRows { get; set; } = 100_000;

        /// <summary>Maximum result columns.</summary>
        public int MaximumColumns { get; set; } = 1_024;

        /// <summary>Maximum result cells.</summary>
        public long MaximumCells { get; set; } = 5_000_000;

        /// <summary>Maximum aggregate characters in column names and string values.</summary>
        public long MaximumCharacters { get; set; } = 64L * 1024 * 1024;
    }

    /// <summary>Detached request supplied to a caller-owned query execution host.</summary>
    public sealed class ExcelQueryExecutionRequest {
        internal ExcelQueryExecutionRequest(
            uint connectionId,
            string connectionName,
            string commandText,
            string worksheetName,
            string tableName,
            bool imported,
            ExcelQueryExecutionPolicy policy) {
            ConnectionId = connectionId;
            ConnectionName = connectionName;
            CommandText = commandText;
            WorksheetName = worksheetName;
            TableName = tableName;
            IsImported = imported;
            Policy = policy;
        }

        /// <summary>Workbook connection identifier.</summary>
        public uint ConnectionId { get; }

        /// <summary>Workbook connection name.</summary>
        public string ConnectionName { get; }

        /// <summary>Opaque command text owned and interpreted by the caller host.</summary>
        public string CommandText { get; }

        /// <summary>Destination worksheet name.</summary>
        public string WorksheetName { get; }

        /// <summary>Destination table name.</summary>
        public string TableName { get; }

        /// <summary>True when the query metadata originated outside the current OfficeIMO session.</summary>
        public bool IsImported { get; }

        /// <summary>Detached execution-policy snapshot. Mutating it does not weaken OfficeIMO's active limits.</summary>
        public ExcelQueryExecutionPolicy Policy { get; }
    }

    /// <summary>Caller-provided query executor. OfficeIMO does not include network or database providers.</summary>
    public interface IExcelQueryExecutionHost {
        /// <summary>Executes a query and returns a bounded, lazily enumerable tabular result.</summary>
        Task<ExcelQueryExecutionResult> ExecuteAsync(
            ExcelQueryExecutionRequest request,
            CancellationToken cancellationToken);
    }

    /// <summary>Tabular rows returned by a caller-owned query host.</summary>
    public sealed class ExcelQueryExecutionResult {
        /// <summary>Creates a tabular query result.</summary>
        public ExcelQueryExecutionResult(
            IReadOnlyList<string> columnNames,
            IEnumerable<IReadOnlyList<object?>> rows) {
            if (columnNames == null) throw new ArgumentNullException(nameof(columnNames));
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            ColumnNames = new ReadOnlyCollection<string>(columnNames.ToArray());
            Rows = rows;
        }

        /// <summary>Ordered result column names.</summary>
        public IReadOnlyList<string> ColumnNames { get; }

        /// <summary>Lazily enumerable rows. OfficeIMO stops enumeration at the configured budget.</summary>
        public IEnumerable<IReadOnlyList<object?>> Rows { get; }
    }

    /// <summary>Options for an OfficeIMO-owned query-backed worksheet table.</summary>
    public sealed class ExcelQueryBackedTableOptions {
        /// <summary>Connection name stored in the workbook.</summary>
        public string ConnectionName { get; set; } = "OfficeIMOQuery";

        /// <summary>Opaque command text passed only to the caller-provided execution host.</summary>
        public string CommandText { get; set; } = string.Empty;

        /// <summary>Destination worksheet.</summary>
        public string WorksheetName { get; set; } = string.Empty;

        /// <summary>Top-left destination cell.</summary>
        public string StartCell { get; set; } = "A1";

        /// <summary>Excel table name.</summary>
        public string TableName { get; set; } = "OfficeIMOQueryTable";

        /// <summary>Initial ordered table schema.</summary>
        public IReadOnlyList<string> ColumnNames { get; set; } = Array.Empty<string>();

        /// <summary>Optional workbook connection description.</summary>
        public string? Description { get; set; }

        /// <summary>Requests refresh-on-open metadata.</summary>
        public bool RefreshOnOpen { get; set; }
    }

    /// <summary>Native query-backed table binding.</summary>
    public sealed class ExcelQueryBackedTableInfo {
        internal ExcelQueryBackedTableInfo(
            uint connectionId,
            string connectionName,
            string commandText,
            string worksheetName,
            string tableName,
            string range,
            bool imported) {
            ConnectionId = connectionId;
            ConnectionName = connectionName;
            CommandText = commandText;
            WorksheetName = worksheetName;
            TableName = tableName;
            Range = range;
            IsImported = imported;
        }

        /// <summary>Workbook connection identifier.</summary>
        public uint ConnectionId { get; }
        /// <summary>Workbook connection name.</summary>
        public string ConnectionName { get; }
        /// <summary>Opaque connection command text.</summary>
        public string CommandText { get; }
        /// <summary>Destination worksheet.</summary>
        public string WorksheetName { get; }
        /// <summary>Destination table.</summary>
        public string TableName { get; }
        /// <summary>Current table range.</summary>
        public string Range { get; }
        /// <summary>True when this binding was loaded rather than authored in this document session.</summary>
        public bool IsImported { get; }
    }

    /// <summary>Completed query refresh details.</summary>
    public sealed class ExcelQueryRefreshResult {
        internal ExcelQueryRefreshResult(ExcelQueryBackedTableInfo source, int rowCount, int columnCount, string range) {
            Source = source;
            RowCount = rowCount;
            ColumnCount = columnCount;
            Range = range;
        }

        /// <summary>Query binding that was refreshed.</summary>
        public ExcelQueryBackedTableInfo Source { get; }
        /// <summary>Written data rows.</summary>
        public int RowCount { get; }
        /// <summary>Written columns.</summary>
        public int ColumnCount { get; }
        /// <summary>Updated table range.</summary>
        public string Range { get; }
    }
}
