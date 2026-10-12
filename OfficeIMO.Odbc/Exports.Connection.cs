using System.Runtime.InteropServices;

namespace OfficeIMO.Odbc {

    internal static unsafe partial class Exports {
        // Windows identifies Unicode drivers by this core entrypoint. Connections use SQLDriverConnect.
        [UnmanagedCallersOnly(EntryPoint = "SQLConnectW")]
        internal static short SQLConnectW(nint id, nint source, short sourceLength, nint user, short userLength, nint password, short passwordLength) =>
            DriverHandles.With<ConnectionHandle>(id, handle => handle.Fail("HYC00", "Use SQLDriverConnect with DRIVER and DBQ; registered DSNs are not supported."));
        [UnmanagedCallersOnly(EntryPoint = "SQLConnect")]
        internal static short SQLConnect(nint id, nint source, short sourceLength, nint user, short userLength, nint password, short passwordLength) =>
            DriverHandles.With<ConnectionHandle>(id, handle => handle.Fail("HYC00", "Use SQLDriverConnect with DRIVER and DBQ; registered DSNs are not supported."));

        [UnmanagedCallersOnly(EntryPoint = "SQLDriverConnectW")]
        internal static short SQLDriverConnectW(nint id, nint window, nint input, short inputLength, nint output, short capacity, short* outputLength, ushort completion) =>
            Connect(id, input, inputLength, output, capacity, outputLength, completion, true);

        [UnmanagedCallersOnly(EntryPoint = "SQLDriverConnect")]
        internal static short SQLDriverConnect(nint id, nint window, nint input, short inputLength, nint output, short capacity, short* outputLength, ushort completion) =>
            Connect(id, input, inputLength, output, capacity, outputLength, completion, false);

        private static short Connect(nint id, nint input, short inputLength, nint output, short capacity, short* outputLength, ushort completion, bool wide) =>
            DriverHandles.With<ConnectionHandle>(id, handle => {
                if (handle.Snapshot is not null) return handle.Fail("08002", "Connection is already open.");
                if (completion != 0) return handle.Fail("HYC00", "Interactive connection dialogs are not supported; use SQL_DRIVER_NOPROMPT.");
                if (capacity < 0) return handle.Fail("HY090", "Negative output buffer length.");
                string connectionString = NativeText.Read(input, inputLength, wide);
                string path = ConnectionSettings.WorkbookPath(connectionString);
                using CancellationTokenSource deadline = new CancellationTokenSource(TimeSpan.FromSeconds(60));
                handle.Snapshot = new WorkbookSnapshot(path, deadline.Token);
                short result = NativeText.Write(connectionString, output, capacity, outputLength, wide);
                if (result == 1) handle.Diagnostics.Add(new("01004", "Connection output was truncated."));
                return result;
            });

        [UnmanagedCallersOnly(EntryPoint = "SQLDisconnect")]
        internal static short SQLDisconnect(nint id) => DriverHandles.With<ConnectionHandle>(id, handle => {
            if (handle.Snapshot is null) return handle.Fail("08003", "Connection is not open.");
            foreach (KeyValuePair<nint, DriverHandle> entry in DriverHandles.All.Where(entry => entry.Value.Parent == id).ToArray()) {
                lock (entry.Value.Sync) ReleaseStatement(entry.Key, (StatementHandle)entry.Value);
            }
            handle.Snapshot.Dispose();
            handle.Snapshot = null;
            return 0;
        });

        [UnmanagedCallersOnly(EntryPoint = "SQLSetConnectAttrW")]
        internal static short SQLSetConnectAttrW(nint id, int attribute, nint value, int length) => SetConnectionAttribute(id, attribute, value);
        [UnmanagedCallersOnly(EntryPoint = "SQLSetConnectAttr")]
        internal static short SQLSetConnectAttr(nint id, int attribute, nint value, int length) => SetConnectionAttribute(id, attribute, value);

        private static short SetConnectionAttribute(nint id, int attribute, nint value) => DriverHandles.With<ConnectionHandle>(id, handle => {
            if (attribute is 101 or 102 && value == 1 || attribute == 103 && value == 0) return 0;
            return handle.Fail("HYC00", "The driver supports read-only access with autocommit enabled.");
        });

        [UnmanagedCallersOnly(EntryPoint = "SQLGetConnectAttrW")]
        internal static short SQLGetConnectAttrW(nint id, int attribute, nint output, int capacity, int* length) => GetConnectionAttribute(id, attribute, output, capacity, length, true);
        [UnmanagedCallersOnly(EntryPoint = "SQLGetConnectAttr")]
        internal static short SQLGetConnectAttr(nint id, int attribute, nint output, int capacity, int* length) => GetConnectionAttribute(id, attribute, output, capacity, length, false);

        private static short GetConnectionAttribute(nint id, int attribute, nint output, int capacity, int* length, bool wide) =>
            DriverHandles.With<ConnectionHandle>(id, handle => {
                if (attribute == 109) {
                    short actual = 0;
                    short result = NativeText.Write(handle.Snapshot is null ? "" : System.IO.Path.GetFileName(handle.Snapshot.Path), output,
                        wide ? capacity / 2 : capacity, &actual, wide);
                    if (length != null) *length = actual * (wide ? 2 : 1);
                    if (result == 1) handle.Diagnostics.Add(new("01004", "Current catalog was truncated."));
                    return result;
                }
                if (output == 0) return handle.Fail("HY009", "Output buffer is required.");
                if (attribute is not (101 or 102 or 103 or 1209)) return handle.Fail("HYC00", "Unsupported connection attribute.");
                *(int*)output = attribute is 101 or 102 ? 1 : attribute == 1209 && handle.Snapshot is null ? 1 : 0;
                if (length != null) *length = sizeof(int);
                return 0;
            });

        [UnmanagedCallersOnly(EntryPoint = "SQLEndTran")]
        internal static short SQLEndTran(short kind, nint id, short completion) => DriverHandles.With<DriverHandle>(id, handle =>
            kind != handle.Kind || kind is not (1 or 2) ? (short)-2 :
                completion is 0 or 1 ? (short)0 : handle.Fail("HY012", "Invalid transaction completion type."));
    }
}
