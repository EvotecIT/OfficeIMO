using System.Runtime.InteropServices;
using System.Data.Common;

namespace OfficeIMO.Odbc {

    internal static unsafe partial class Exports {
        [UnmanagedCallersOnly(EntryPoint = "SQLPrepareW")]
        internal static short SQLPrepareW(nint id, nint sql, int length) => Prepare(id, sql, length, true);
        [UnmanagedCallersOnly(EntryPoint = "SQLPrepare")]
        internal static short SQLPrepare(nint id, nint sql, int length) => Prepare(id, sql, length, false);
        private static short Prepare(nint id, nint sql, int length, bool wide) => DriverHandles.With<StatementHandle>(id, statement => {
            statement.CloseCursor();
            statement.Sql = null;
            statement.ParameterCount = 0;
            statement.Sql = QueryEnvelope.Prepare(NativeText.Read(sql, length, wide), out int parameterCount);
            statement.ParameterCount = parameterCount;
            return 0;
        });

        [UnmanagedCallersOnly(EntryPoint = "SQLExecDirectW")]
        internal static short SQLExecDirectW(nint id, nint sql, int length) => Direct(id, sql, length, true);
        [UnmanagedCallersOnly(EntryPoint = "SQLExecDirect")]
        internal static short SQLExecDirect(nint id, nint sql, int length) => Direct(id, sql, length, false);
        private static short Direct(nint id, nint sql, int length, bool wide) => DriverHandles.With<StatementHandle>(id, statement => {
            statement.CloseCursor();
            statement.Sql = null;
            statement.ParameterCount = 0;
            statement.Sql = QueryEnvelope.Prepare(NativeText.Read(sql, length, wide), out int parameterCount);
            statement.ParameterCount = parameterCount;
            return Execute(statement);
        });

        [UnmanagedCallersOnly(EntryPoint = "SQLExecute")]
        internal static short SQLExecute(nint id) => DriverHandles.With<StatementHandle>(id, Execute);

        private static short Execute(StatementHandle statement) {
            statement.CloseCursor();
            if (statement.Sql is null) return statement.Fail("HY010", "Prepare a SELECT statement first.");
            if (!DriverHandles.All.TryGetValue(statement.Parent, out DriverHandle? parent) || parent is not ConnectionHandle { Snapshot: { } snapshot })
                return statement.Fail("08003", "Connection is not open.");
            Dictionary<string, object?> parameters = new Dictionary<string, object?>();
            for (ushort i = 1; i <= statement.ParameterCount; i++) {
                if (!statement.Parameters.TryGetValue(i, out ParameterBinding? binding)) {
                    return statement.Fail("07002", "Bind every parameter before execution.");
                } else parameters["@p" + i] = NativeValue.Read(binding);
            }
            statement.Cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(statement.TimeoutSeconds));
            try {
                statement.Reader = snapshot.Query(statement.Sql, parameters, statement.Cancellation.Token);
                statement.Cancellation.CancelAfter(Timeout.Infinite);
                if (statement.Reader.FieldCount > 256) {
                    statement.CloseCursor();
                    return statement.Fail("HY001", "Result exceeds the 256 column limit.");
                }
                return 0;
            } catch { statement.CloseCursor(); throw; }
        }

        private static short EnsureMetadata(StatementHandle statement) {
            if (statement.Reader is not null || statement.PreparedColumns is not null) return 0;
            if (statement.Sql is null) return statement.Fail("HY010", "Prepare a SELECT statement first.");
            statement.Cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(statement.TimeoutSeconds));
            try {
                IReadOnlyList<DBAClientX.SQLiteQueryColumn> columns = Snapshot(statement).QueryColumns(statement.Sql, statement.Cancellation.Token);
                if (columns.Count > 256) return statement.Fail("HY001", "Result exceeds the 256 column limit.");
                statement.PreparedColumns = columns;
                return 0;
            } finally {
                statement.Cancellation.Dispose();
                statement.Cancellation = null;
            }
        }

        [UnmanagedCallersOnly(EntryPoint = "SQLFetch")]
        internal static short SQLFetch(nint id) => DriverHandles.With<StatementHandle>(id, Fetch);

        [UnmanagedCallersOnly(EntryPoint = "SQLFetchScroll")]
        internal static short SQLFetchScroll(nint id, short orientation, nint offset) => DriverHandles.With<StatementHandle>(id, statement =>
            orientation == 1 ? Fetch(statement) : statement.Fail("HY106", "Only forward-only SQL_FETCH_NEXT is supported."));

        private static short Fetch(StatementHandle statement) {
            statement.OnRow = false;
            statement.DataColumn = 0;
            statement.DataOffset = 0;
            if (statement.RowsFetchedPointer != 0) *(nuint*)statement.RowsFetchedPointer = 0;
            if (statement.RowStatusPointer != 0) *(ushort*)statement.RowStatusPointer = 3;
            if (statement.Reader is not { } reader) return statement.Fail("24000", "No result cursor is open.");
            statement.Cancellation?.Token.ThrowIfCancellationRequested();
            if (statement.MaximumRows > 0 && statement.Fetched >= statement.MaximumRows) return 100;
            bool hasRow;
            statement.Cancellation?.CancelAfter(TimeSpan.FromSeconds(statement.TimeoutSeconds));
            try {
                hasRow = reader.Read();
                statement.Cancellation?.Token.ThrowIfCancellationRequested();
            } finally { statement.Cancellation?.CancelAfter(Timeout.Infinite); }
            if (!hasRow) return 100;
            if (++statement.Fetched > 100_000) return statement.Fail("HY001", "Result exceeds the 100,000 row limit.");
            statement.OnRow = true;
            if (statement.RowsFetchedPointer != 0) *(nuint*)statement.RowsFetchedPointer = 1;
            if (statement.RowStatusPointer != 0) *(ushort*)statement.RowStatusPointer = 0;
            short status = 0;
            foreach (KeyValuePair<ushort, ColumnBinding> binding in statement.Bindings) {
                short result = NativeValue.Write(statement, binding.Key, binding.Value.Type, binding.Value.Buffer,
                    binding.Value.Length, binding.Value.Indicator, false);
                if (result == -1) { if (statement.RowStatusPointer != 0) *(ushort*)statement.RowStatusPointer = 5; return result; }
                if (result == 1) status = 1;
            }
            if (status == 1 && statement.RowStatusPointer != 0) *(ushort*)statement.RowStatusPointer = 6;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "SQLBindCol")]
        internal static short SQLBindCol(nint id, ushort column, short type, nint buffer, nint length, nint indicator) =>
            DriverHandles.With<StatementHandle>(id, statement => {
                if (column == 0 || column > 256) return statement.Fail("07009", "Invalid column ordinal.");
                if (length < 0) return statement.Fail("HY090", "Negative buffer length.");
                if (buffer == 0) statement.Bindings.Remove(column);
                else statement.Bindings[column] = new(type, buffer, length, indicator);
                return 0;
            });

        [UnmanagedCallersOnly(EntryPoint = "SQLBindParameter")]
        internal static short SQLBindParameter(nint id, ushort ordinal, short direction, short type, short sqlType, nuint size,
            short scale, nint buffer, nint length, nint indicator) => DriverHandles.With<StatementHandle>(id, statement => {
                if (direction != 1) return statement.Fail("HYC00", "Only input parameters are supported.");
                if (ordinal == 0 || ordinal > 256) return statement.Fail("07009", "Invalid parameter ordinal.");
                if (length < 0) return statement.Fail("HY090", "Negative parameter buffer length.");
                short defaultType = NativeParameter.DefaultCType(sqlType);
                if (defaultType == 0) return statement.Fail("HY004", "Unsupported SQL parameter type.");
                short effectiveType = type == 99 ? defaultType : type;
                if (!NativeParameter.Supports(effectiveType, sqlType))
                    return statement.Fail("07006", "Bind a matching C and SQL parameter type; cross-type conversions are not supported.");
                if (sqlType == 93 && scale is < 0 or > 7) return statement.Fail("HY104", "Timestamp parameters support zero through seven fractional digits.");
                statement.Parameters[ordinal] = new(effectiveType, buffer, length, indicator, sqlType, size, scale);
                return 0;
            });

        [UnmanagedCallersOnly(EntryPoint = "SQLGetData")]
        internal static short SQLGetData(nint id, ushort column, short type, nint buffer, nint capacity, nint indicator) =>
            DriverHandles.With<StatementHandle>(id, statement => NativeValue.Write(statement, column, type, buffer, capacity, indicator, true));

        [UnmanagedCallersOnly(EntryPoint = "SQLNumResultCols")]
        internal static short SQLNumResultCols(nint id, short* count) => DriverHandles.With<StatementHandle>(id, statement => {
            if (count == null) return statement.Fail("HY009", "Output buffer is required.");
            short result = EnsureMetadata(statement);
            if (result == 0) *count = checked((short)statement.ColumnCount);
            return result;
        });

        [UnmanagedCallersOnly(EntryPoint = "SQLNumParams")]
        internal static short SQLNumParams(nint id, short* count) => DriverHandles.With<StatementHandle>(id, statement => {
            if (count == null) return statement.Fail("HY009", "Output buffer is required.");
            *count = checked((short)statement.ParameterCount);
            return 0;
        });

        [UnmanagedCallersOnly(EntryPoint = "SQLRowCount")]
        internal static short SQLRowCount(nint id, nint* count) => DriverHandles.With<StatementHandle>(id, statement => {
            if (count == null) return statement.Fail("HY009", "Output buffer is required.");
            *count = -1;
            return 0;
        });

        [UnmanagedCallersOnly(EntryPoint = "SQLCloseCursor")]
        internal static short SQLCloseCursor(nint id) => DriverHandles.With<StatementHandle>(id, statement => {
            if (statement.Reader is null) return statement.Fail("24000", "No result cursor is open.");
            statement.CloseCursor(); return 0;
        });
        [UnmanagedCallersOnly(EntryPoint = "SQLMoreResults")]
        internal static short SQLMoreResults(nint id) => DriverHandles.With<StatementHandle>(id, statement => { statement.CloseCursor(); return 100; });
        [UnmanagedCallersOnly(EntryPoint = "SQLFreeStmt")]
        internal static short SQLFreeStmt(nint id, ushort option) => DriverHandles.With<StatementHandle>(id, statement => {
            switch (option) {
                case 0: statement.CloseCursor(); return 0;
                case 2: statement.Bindings.Clear(); return 0;
                case 3: statement.Parameters.Clear(); return 0;
                default: return statement.Fail("HYC00", "Use SQLFreeHandle to release statement handles.");
            }
        });

        [UnmanagedCallersOnly(EntryPoint = "SQLCancel")]
        internal static short SQLCancel(nint id) {
            if (!DriverHandles.All.TryGetValue(id, out DriverHandle? handle) || handle is not StatementHandle statement || handle.Freed) return -2;
            try { statement.Cancellation?.Cancel(); return 0; }
            catch (ObjectDisposedException) { return 0; }
            catch { return -1; }
        }
    }
}
