using System.Runtime.InteropServices;

namespace OfficeIMO.Odbc {

    internal static unsafe partial class Exports {
        [UnmanagedCallersOnly(EntryPoint = "SQLSetStmtAttrW")]
        internal static short SQLSetStmtAttrW(nint id, int attribute, nint value, int length) => SetStatementAttribute(id, attribute, value);
        [UnmanagedCallersOnly(EntryPoint = "SQLSetStmtAttr")]
        internal static short SQLSetStmtAttr(nint id, int attribute, nint value, int length) => SetStatementAttribute(id, attribute, value);

        private static short SetStatementAttribute(nint id, int attribute, nint value) => DriverHandles.With<StatementHandle>(id, statement => {
            switch (attribute) {
                case 0:
                    if (value < 0 || value > 300) return statement.Fail("HY024", "Query timeout must be between 0 and 300 seconds.");
                    statement.TimeoutSeconds = value == 0 ? 30 : (int)value;
                    if (value == 0) { statement.Diagnostics.Add(new("01S02", "A zero timeout uses the bounded 30 second default.")); return 1; }
                    return 0;
                case 1:
                    if (value < 0 || value > 100_000) return statement.Fail("HY024", "Maximum rows must be between 0 and 100,000.");
                    statement.MaximumRows = (long)value; return 0;
                case 25: statement.RowStatusPointer = value; return 0;
                case 26: statement.RowsFetchedPointer = value; return 0;
                case 2:
                    if (value == 1) return 0;
                    return statement.Fail("HYC00", "ODBC escape translation is not supported; use SQL_NOSCAN_ON and SQLite SELECT syntax.");
                case 4: case 5: case 6: case 18: case 20:
                    if (value == 0) return 0;
                    break;
                case 7: case 11: case 22: case 27:
                    if (value == 1) return 0;
                    break;
            }
            return statement.Fail("HYC00", "Only scalar, read-only, forward-only statement attributes are supported.");
        });

        [UnmanagedCallersOnly(EntryPoint = "SQLGetStmtAttrW")]
        internal static short SQLGetStmtAttrW(nint id, int attribute, nint output, int capacity, int* length) => GetStatementAttribute(id, attribute, output, length);
        [UnmanagedCallersOnly(EntryPoint = "SQLGetStmtAttr")]
        internal static short SQLGetStmtAttr(nint id, int attribute, nint output, int capacity, int* length) => GetStatementAttribute(id, attribute, output, length);

        private static short GetStatementAttribute(nint id, int attribute, nint output, int* length) => DriverHandles.With<StatementHandle>(id, statement => {
            if (output == 0) return statement.Fail("HY009", "Output buffer is required.");
            nint value;
            switch (attribute) {
                case 0: value = statement.TimeoutSeconds; break;
                case 1: value = (nint)statement.MaximumRows; break;
                case 25: value = statement.RowStatusPointer; break;
                case 26: value = statement.RowsFetchedPointer; break;
                case 10010: case 10011: case 10012: case 10013: value = statement.Descriptors[attribute]; break;
                case 4: case 5: case 6: case 18: case 20: value = 0; break;
                case 2: case 7: case 11: case 22: case 27: value = 1; break;
                default: return statement.Fail("HYC00", "Unsupported statement attribute.");
            }
            *(nint*)output = value;
            if (length != null) *length = sizeof(nint);
            return 0;
        });

        [UnmanagedCallersOnly(EntryPoint = "SQLGetDiagFieldW")]
        internal static short SQLGetDiagFieldW(short kind, nint id, short record, short field, nint output, short capacity, short* length) =>
            DiagnosticField(kind, id, record, field, output, capacity, length, true);
        [UnmanagedCallersOnly(EntryPoint = "SQLGetDiagField")]
        internal static short SQLGetDiagField(short kind, nint id, short record, short field, nint output, short capacity, short* length) =>
            DiagnosticField(kind, id, record, field, output, capacity, length, false);

        private static short DiagnosticField(short kind, nint id, short record, short field, nint output, short capacity, short* length, bool wide) {
            if (!DriverHandles.All.TryGetValue(id, out DriverHandle? handle) || handle.Kind != kind) return -2;
            lock (handle.Sync) {
                try {
                    if (handle.Freed) return -2;
                    if (capacity < 0) return -1;
                    if (field == 2 && output != 0) { *(int*)output = handle.Diagnostics.Count; return 0; }
                    // SQL_DIAG_RETURNCODE is maintained by the Driver Manager.
                    if (field == 1) return -1;
                    if (field == 3 && output != 0 && handle is StatementHandle) { *(nint*)output = -1; return 0; }
                    if (record < 1 || record > handle.Diagnostics.Count) return 100;
                    DriverDiagnostic diagnostic = handle.Diagnostics[record - 1];
                    if (field == 5 && output != 0) { *(int*)output = 0; return 0; }
                    if (field == -1248 && output != 0) { *(nint*)output = -1; return 0; }
                    if (field == -1247 && output != 0) { *(int*)output = -1; return 0; }
                    string? value = field switch { 4 => diagnostic.State, 6 => diagnostic.Message, 8 or 9 => "ISO 9075", _ => null };
                    if (value is null) return -1;
                    short units = 0;
                    short result = NativeText.Write(value, output, wide ? capacity / 2 : capacity, &units, wide);
                    if (length != null) *length = checked((short)(units * (wide ? 2 : 1)));
                    return result;
                } catch { return -1; }
            }
        }
    }
}
