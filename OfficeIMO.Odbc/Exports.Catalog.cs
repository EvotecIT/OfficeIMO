using System.Data;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace OfficeIMO.Odbc {

    internal static unsafe partial class Exports {
        [UnmanagedCallersOnly(EntryPoint = "SQLTablesW")]
        internal static short SQLTablesW(nint id, nint catalog, short catalogLength, nint schema, short schemaLength, nint table, short tableLength, nint types, short typesLength) =>
            Tables(id, catalog, catalogLength, schema, schemaLength, table, tableLength, types, typesLength, true);
        [UnmanagedCallersOnly(EntryPoint = "SQLTables")]
        internal static short SQLTables(nint id, nint catalog, short catalogLength, nint schema, short schemaLength, nint table, short tableLength, nint types, short typesLength) =>
            Tables(id, catalog, catalogLength, schema, schemaLength, table, tableLength, types, typesLength, false);

        private static short Tables(nint id, nint catalog, short catalogLength, nint schema, short schemaLength, nint table, short tableLength, nint types, short typesLength, bool wide) =>
            DriverHandles.With<StatementHandle>(id, statement => {
                WorkbookSnapshot snapshot = Snapshot(statement);
                string? cat = Pattern(catalog, catalogLength, wide), owner = Pattern(schema, schemaLength, wide);
                string? name = Pattern(table, tableLength, wide), kind = Pattern(types, typesLength, wide);
                DataTable result = NewMetadata([("TABLE_CAT", typeof(string)), ("TABLE_SCHEM", typeof(string)),
                    ("TABLE_NAME", typeof(string)), ("TABLE_TYPE", typeof(string)), ("REMARKS", typeof(string))]);
                if (cat == "%" && owner == "" && name == "") result.Rows.Add("", DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value);
                else if (owner == "%" && cat == "" && name == "") result.Rows.Add(DBNull.Value, "", DBNull.Value, DBNull.Value, DBNull.Value);
                else if (kind == "%" && cat == "" && owner == "" && name == "") result.Rows.Add(DBNull.Value, DBNull.Value, DBNull.Value, "TABLE", DBNull.Value);
                else if (Matches("", cat) && Matches("", owner) && (string.IsNullOrEmpty(kind) || kind.Split(',').Any(value => value.Trim(' ', '\'').Equals("TABLE", StringComparison.OrdinalIgnoreCase)))) {
                    foreach (SnapshotTable item in snapshot.Tables.Where(item => Matches(item.Name, name)).OrderBy(item => item.Name, StringComparer.Ordinal)) result.Rows.Add(DBNull.Value, DBNull.Value, item.Name, "TABLE", "Read-only worksheet snapshot");
                }
                SetMetadata(statement, result);
                return 0;
            });

        [UnmanagedCallersOnly(EntryPoint = "SQLColumnsW")]
        internal static short SQLColumnsW(nint id, nint catalog, short catalogLength, nint schema, short schemaLength, nint table, short tableLength, nint column, short columnLength) =>
            Columns(id, catalog, catalogLength, schema, schemaLength, table, tableLength, column, columnLength, true);
        [UnmanagedCallersOnly(EntryPoint = "SQLColumns")]
        internal static short SQLColumns(nint id, nint catalog, short catalogLength, nint schema, short schemaLength, nint table, short tableLength, nint column, short columnLength) =>
            Columns(id, catalog, catalogLength, schema, schemaLength, table, tableLength, column, columnLength, false);

        private static short Columns(nint id, nint catalog, short catalogLength, nint schema, short schemaLength, nint table, short tableLength, nint column, short columnLength, bool wide) =>
            DriverHandles.With<StatementHandle>(id, statement => {
                WorkbookSnapshot snapshot = Snapshot(statement);
                string? cat = Pattern(catalog, catalogLength, wide), owner = Pattern(schema, schemaLength, wide);
                string? name = Pattern(table, tableLength, wide), field = Pattern(column, columnLength, wide);
                DataTable result = NewMetadata([
                    ("TABLE_CAT", typeof(string)), ("TABLE_SCHEM", typeof(string)), ("TABLE_NAME", typeof(string)), ("COLUMN_NAME", typeof(string)),
                    ("DATA_TYPE", typeof(short)), ("TYPE_NAME", typeof(string)), ("COLUMN_SIZE", typeof(int)), ("BUFFER_LENGTH", typeof(int)),
                    ("DECIMAL_DIGITS", typeof(short)), ("NUM_PREC_RADIX", typeof(short)), ("NULLABLE", typeof(short)), ("REMARKS", typeof(string)),
                    ("COLUMN_DEF", typeof(string)), ("SQL_DATA_TYPE", typeof(short)), ("SQL_DATETIME_SUB", typeof(short)),
                    ("CHAR_OCTET_LENGTH", typeof(int)), ("ORDINAL_POSITION", typeof(int)), ("IS_NULLABLE", typeof(string))]);
                if (Matches("", cat) && Matches("", owner)) foreach (SnapshotTable item in snapshot.Tables.Where(item => Matches(item.Name, name)).OrderBy(item => item.Name, StringComparer.Ordinal)) {
                    for (int i = 0; i < item.Columns.Length; i++) {
                        SnapshotColumn definition = item.Columns[i];
                        if (!Matches(definition.Name, field)) continue;
                        short type = definition.Type == typeof(double) ? (short)8 : definition.Type == typeof(bool) ? (short)-7 : definition.Type == typeof(DateTime) ? (short)93 : (short)-9;
                        result.Rows.Add(DBNull.Value, DBNull.Value, item.Name, definition.Name, type, NativeValue.TypeName(type),
                            (int)NativeValue.CatalogColumnSize(type), (int)NativeValue.CatalogOctetLength(type), type == 8 ? DBNull.Value : (object)(short)(type == 93 ? 7 : 0),
                            type is 8 or -7 ? (object)(short)10 : DBNull.Value, (short)1, DBNull.Value, DBNull.Value,
                            (short)(type == 93 ? 9 : type), type == 93 ? (object)(short)3 : DBNull.Value,
                            type == -9 ? (object)(64 * 1024 * 1024) : DBNull.Value, i + 1, "YES");
                    }
                }
                SetMetadata(statement, result);
                return 0;
            });

        [UnmanagedCallersOnly(EntryPoint = "SQLGetTypeInfoW")]
        internal static short SQLGetTypeInfoW(nint id, short type) => TypeInfo(id, type);
        [UnmanagedCallersOnly(EntryPoint = "SQLGetTypeInfo")]
        internal static short SQLGetTypeInfo(nint id, short type) => TypeInfo(id, type);

        private static short TypeInfo(nint id, short selected) => DriverHandles.With<StatementHandle>(id, statement => {
            Snapshot(statement);
            DataTable result = NewMetadata([
                ("TYPE_NAME", typeof(string)), ("DATA_TYPE", typeof(short)), ("COLUMN_SIZE", typeof(int)), ("LITERAL_PREFIX", typeof(string)),
                ("LITERAL_SUFFIX", typeof(string)), ("CREATE_PARAMS", typeof(string)), ("NULLABLE", typeof(short)), ("CASE_SENSITIVE", typeof(short)),
                ("SEARCHABLE", typeof(short)), ("UNSIGNED_ATTRIBUTE", typeof(short)), ("FIXED_PREC_SCALE", typeof(short)), ("AUTO_UNIQUE_VALUE", typeof(short)),
                ("LOCAL_TYPE_NAME", typeof(string)), ("MINIMUM_SCALE", typeof(short)), ("MAXIMUM_SCALE", typeof(short)), ("SQL_DATA_TYPE", typeof(short)),
                ("SQL_DATETIME_SUB", typeof(short)), ("NUM_PREC_RADIX", typeof(int)), ("INTERVAL_PRECISION", typeof(short))]);
            foreach (short type in new short[] { -9, -7, -5, 4, 5, 8, 93 }) {
                if (selected != 0 && selected != type) continue;
                result.Rows.Add(NativeValue.TypeName(type), type, type == 8 ? 53 : (int)NativeValue.CatalogColumnSize(type), type is -9 or 93 ? "'" : DBNull.Value,
                    type is -9 or 93 ? "'" : DBNull.Value, DBNull.Value, (short)1, (short)(type == -9 ? 1 : 0), (short)3,
                    type is -9 or 93 ? DBNull.Value : (short)0, (short)0, (short)0, NativeValue.TypeName(type), (short)0,
                    (short)(type == 93 ? 7 : 0), (short)(type == 93 ? 9 : type), type == 93 ? (object)(short)3 : DBNull.Value,
                    type is -9 or 93 ? DBNull.Value : type == 8 ? 2 : 10, DBNull.Value);
            }
            SetMetadata(statement, result);
            return 0;
        });

        private static WorkbookSnapshot Snapshot(StatementHandle statement) =>
            DriverHandles.All.TryGetValue(statement.Parent, out DriverHandle? parent) && parent is ConnectionHandle { Snapshot: { } snapshot }
                ? snapshot : throw new InvalidOperationException("Connection is not open.");

        private static DataTable NewMetadata((string Name, Type Type)[] columns) {
            DataTable result = new DataTable();
            foreach ((string Name, Type Type) column in columns) {
                // Literal types keep metadata creation trim-safe in the native library.
                if (column.Type == typeof(short)) result.Columns.Add(column.Name, typeof(short));
                else if (column.Type == typeof(int)) result.Columns.Add(column.Name, typeof(int));
                else result.Columns.Add(column.Name, typeof(string));
            }
            return result;
        }

        private static void SetMetadata(StatementHandle statement, DataTable result) {
            statement.CloseCursor();
            statement.Reader = result.CreateDataReader();
        }

        private static string? Pattern(nint pointer, short length, bool wide) => pointer == 0 ? null : NativeText.Read(pointer, length, wide);

        private static bool Matches(string value, string? pattern) {
            if (pattern is null) return true;
            System.Text.StringBuilder expression = new System.Text.StringBuilder("\\A");
            for (int i = 0; i < pattern.Length; i++) {
                char character = pattern[i];
                if (character == '\\' && i + 1 < pattern.Length) expression.Append(Regex.Escape(pattern[++i].ToString()));
                else expression.Append(character == '%' ? ".*" : character == '_' ? "." : Regex.Escape(character.ToString()));
            }
            return Regex.IsMatch(value, expression.Append("\\z").ToString(), RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline | RegexOptions.NonBacktracking,
                TimeSpan.FromSeconds(1));
        }
    }
}
