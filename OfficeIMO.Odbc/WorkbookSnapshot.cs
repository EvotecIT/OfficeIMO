using System.Data;
using System.Data.Common;
using System.Globalization;
using DBAClientX;
using OfficeIMO.Excel;

namespace OfficeIMO.Odbc {

    /// <summary>An immutable, bounded workbook projection. DbaClientX owns every provider connection.</summary>
    internal sealed class WorkbookSnapshot : IDisposable {
        private readonly SQLite _writer = new();
        private readonly SQLite _queries;
        private readonly DbConnection _anchor;
        private readonly string _connectionString;
        private readonly string _database;
        internal IReadOnlyList<SnapshotTable> Tables { get; }
        internal string Path { get; }
        internal string CatalogName => System.IO.Path.GetFileName(Path);

        internal WorkbookSnapshot(string path, CancellationToken cancellationToken) {
            Path = System.IO.Path.GetFullPath(path);
            cancellationToken.ThrowIfCancellationRequested();
            ConnectionSettings.ValidateStoragePath(Path);
            // A random named in-memory database lets independently owned query readers share one snapshot.
            string database = "file:officeimo-" + Guid.NewGuid().ToString("N") + "?mode=memory&cache=shared";
            _database = database;
            _anchor = _writer.OpenDbConnection(database, new SQLiteConnectionOptions {
                Pooling = false, EnableWriteAheadLogging = false, UseNormalSynchronousMode = false
            });
            _connectionString = _anchor.ConnectionString;
            _queries = new SQLite {
                MaxRetryAttempts = 1,
                ConnectionOptions = new SQLiteConnectionOptions { Pooling = false, EnableWriteAheadLogging = false, UseNormalSynchronousMode = false },
                ConfigureConnection = connection => {
                    using DbCommand command = connection.CreateCommand();
                    command.CommandText = "PRAGMA query_only=ON; PRAGMA trusted_schema=OFF; PRAGMA temp_store=MEMORY;";
                    command.ExecuteNonQuery();
                }
            };
            try {
                Tables = Import(cancellationToken);
            } catch {
                Dispose();
                throw;
            }
        }

        private List<SnapshotTable> Import(CancellationToken cancellationToken) {
            // Buffer the source once, so all worksheets belong to the same file snapshot.
            cancellationToken.ThrowIfCancellationRequested();
            using FileStream input = new FileStream(Path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, FileOptions.Asynchronous | FileOptions.SequentialScan);
            const int maxInputBytes = 64 * 1024 * 1024;
            if (input.Length > maxInputBytes) throw new InvalidDataException("Workbook exceeds the 64 MiB input limit.");
            byte[] bytes = new byte[checked((int)input.Length)];
            input.ReadExactlyAsync(bytes, cancellationToken).AsTask().GetAwaiter().GetResult();
            ExcelReadOptions options = new ExcelReadOptions {
                MaxInputBytes = maxInputBytes, MaxWorksheets = 64, CancellationToken = cancellationToken,
                HasHeaderRow = true, InferSchema = false, MaxDataReaderColumns = 256,
                MaxSharedStringCharacters = 32L * 1024 * 1024, MaxSharedStringItemCharacters = 32_767
            };
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(bytes, options);
            List<SnapshotTable> tables = new List<SnapshotTable>();
            HashSet<string> names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            long cells = 0, characters = 0;
            int rows = 0;
            do {
                cancellationToken.ThrowIfCancellationRequested();
                if (reader.FieldCount == 0) continue;
                if (reader.FieldCount > 256) throw new InvalidDataException("A worksheet exceeds the 256 column limit.");
                string name = reader.CurrentSheetName;
                if (name.StartsWith("sqlite_", StringComparison.OrdinalIgnoreCase) || !names.Add(name))
                    throw new InvalidDataException("Worksheet names must be distinct SQLite table identifiers.");
                string[] columns = Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).ToArray();
                if (columns.Any(column => column.Length > 255)) throw new InvalidDataException("Worksheet headers exceed 255 characters.");
                if (columns.Distinct(StringComparer.OrdinalIgnoreCase).Count() != columns.Length)
                    throw new InvalidDataException("Worksheet headers must be distinct, ignoring case.");
                Type?[] types = new Type?[columns.Length];
                List<object?[]> values = new List<object?[]>();
                while (reader.Read()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (++rows > 100_000 || (cells += columns.Length) > 1_000_000)
                        throw new InvalidDataException("Workbook exceeds 100,000 rows or 1,000,000 cells.");
                    object?[] row = new object?[columns.Length];
                    for (int i = 0; i < row.Length; i++) {
                        object value = reader.GetValue(i);
                        if (value is DBNull || value is null) continue;
                        Type type = value is bool ? typeof(bool) : value is DateTime ? typeof(DateTime) :
                            value is double or float or decimal or int or long or short or byte ? typeof(double) : typeof(string);
                        types[i] = types[i] is null ? type : types[i] == type ? type : typeof(string);
                        if (value is string text && (characters += text.Length) > 32 * 1024 * 1024)
                            throw new InvalidDataException("Workbook exceeds the 32 Mi character text limit.");
                        row[i] = value;
                    }
                    values.Add(row);
                }
                SnapshotColumn[] metadata = columns.Select((column, i) => new SnapshotColumn(column, types[i] ?? typeof(string))).ToArray();
                using DbTransaction transaction = _anchor.BeginTransaction();
                using DbCommand create = _anchor.CreateCommand();
                create.Transaction = transaction;
                create.CommandText = "CREATE TABLE " + Quote(name) + " (" + string.Join(",", metadata.Select(column =>
                    Quote(column.Name) + " " + column.StorageType)) + ")";
                create.ExecuteNonQuery();
                using DbCommand insert = _anchor.CreateCommand();
                insert.Transaction = transaction;
                insert.CommandText = "INSERT INTO " + Quote(name) + " VALUES (" +
                    string.Join(",", Enumerable.Range(0, columns.Length).Select(i => "@p" + i)) + ")";
                for (int i = 0; i < columns.Length; i++) {
                    DbParameter parameter = insert.CreateParameter();
                    parameter.ParameterName = "@p" + i;
                    insert.Parameters.Add(parameter);
                }
                insert.Prepare();
                foreach (object?[] row in values) {
                    cancellationToken.ThrowIfCancellationRequested();
                    for (int i = 0; i < row.Length; i++) insert.Parameters[i].Value = Normalize(row[i], metadata[i].Type);
                    insert.ExecuteNonQuery();
                }
                transaction.Commit();
                tables.Add(new SnapshotTable(name, metadata));
            } while (reader.NextResult());
            return tables;
        }

        internal DbDataReader Query(string sql, IDictionary<string, object?> parameters, CancellationToken cancellationToken) {
            QueryEnvelope.Validate(sql);
            return _queries.QueryReaderWithConnectionStringAsync(_connectionString, sql, parameters,
                cancellationToken: cancellationToken).GetAwaiter().GetResult();
        }

        internal IReadOnlyList<SQLiteQueryColumn> QueryColumns(string sql, CancellationToken cancellationToken) {
            QueryEnvelope.Validate(sql);
            return _queries.GetQueryColumns(_database, sql, cancellationToken);
        }

        private static object Normalize(object? value, Type type) => value is null ? DBNull.Value :
            type == typeof(string) ? value is DateTime date ? date.ToString("O", CultureInfo.InvariantCulture) :
                Convert.ToString(value, CultureInfo.InvariantCulture)! :
            type == typeof(double) ? Convert.ToDouble(value, CultureInfo.InvariantCulture) : value;

        internal static string Quote(string name) => "\"" + name.Replace("\"", "\"\"") + "\"";

        public void Dispose() {
            _queries.Dispose();
            _anchor.Dispose();
            _writer.Dispose();
        }
    }

    internal sealed record SnapshotTable(string Name, SnapshotColumn[] Columns);
    internal sealed record SnapshotColumn(string Name, Type Type) {
        internal string StorageType => Type == typeof(double) ? "REAL" : Type == typeof(bool) ? "BOOLEAN" :
            Type == typeof(DateTime) ? "DATETIME" : "TEXT";
    }
}
