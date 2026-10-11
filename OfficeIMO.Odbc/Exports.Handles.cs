using System.Runtime.InteropServices;

namespace OfficeIMO.Odbc;

internal static unsafe partial class Exports {
    // Native AOT cannot be unloaded. Pin this module before returning the first driver handle.
    [DllImport("kernel32.dll", EntryPoint = "GetModuleHandleExW", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PinModule(uint flags, nint address, out nint module);

    [DllImport("kernel32.dll", EntryPoint = "GetModuleFileNameW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint ModuleFileName(nint module, char* path, uint capacity);

    private static bool _runtimeInitialized;

    [UnmanagedCallersOnly(EntryPoint = "SQLAllocHandle")]
    internal static short SQLAllocHandle(short kind, nint parent, nint* output) {
        try {
            if (output == null) return -1;
            *output = 0;
            if (!OperatingSystem.IsWindows() || !Environment.Is64BitProcess) return -1;
            nint address = (nint)(delegate* unmanaged<short, nint, nint*, short>)&SQLAllocHandle;
            if (!PinModule(5, address, out nint module)) return -1;
            lock (DriverHandles.All) {
                if (!_runtimeInitialized) {
                    char* path = stackalloc char[32768];
                    uint length = ModuleFileName(module, path, 32768);
                    if (length == 0 || length >= 32768) return -1;
                    DBAClientX.SQLite.InitializeNativeLibraryDirectory(System.IO.Path.GetDirectoryName(new string(path, 0, (int)length))!);
                    _runtimeInitialized = true;
                }
            }
            short result = DriverHandles.Allocate(kind, parent, out nint id);
            *output = id;
            return result;
        } catch { return -1; }
    }

    [UnmanagedCallersOnly(EntryPoint = "SQLFreeHandle")]
    internal static short SQLFreeHandle(short kind, nint id) => DriverHandles.With<DriverHandle>(id, handle => {
        if (handle.Kind != kind) return -2;
        lock (DriverHandles.All) {
            if (handle is DescriptorHandle) return handle.Fail("HY017", "Implicit descriptors belong to their statement.");
            if (DriverHandles.All.Values.Any(child => child.Parent == id && child is not DescriptorHandle)) return handle.Fail("HY010", "Free child handles first.");
            if (handle is StatementHandle statement) {
                statement.CloseCursor();
                foreach (nint descriptor in statement.Descriptors.Values) {
                    if (DriverHandles.All.TryRemove(descriptor, out var owned)) owned.Freed = true;
                }
            }
            if (handle is ConnectionHandle connection) { connection.Snapshot?.Dispose(); connection.Snapshot = null; }
            handle.Freed = true;
            DriverHandles.All.TryRemove(id, out _);
        }
        return 0;
    });

    [UnmanagedCallersOnly(EntryPoint = "SQLSetEnvAttr")]
    internal static short SQLSetEnvAttr(nint id, int attribute, nint value, int length) =>
        DriverHandles.With<EnvironmentHandle>(id, handle =>
            attribute == 200 && (value == 3 || value == 380) || attribute == 10001 && value == 1 ?
                (short)0 : handle.Fail("HYC00", "Only ODBC 3 and terminated output strings are supported."));

    [UnmanagedCallersOnly(EntryPoint = "SQLGetEnvAttr")]
    internal static short SQLGetEnvAttr(nint id, int attribute, nint output, int capacity, int* length) =>
        DriverHandles.With<EnvironmentHandle>(id, handle => {
            if (output == 0) return handle.Fail("HY009", "Output buffer is required.");
            if (attribute is not (200 or 10001)) return handle.Fail("HYC00", "Unsupported environment attribute.");
            *(int*)output = attribute == 200 ? 3 : 1;
            if (length != null) *length = sizeof(int);
            return 0;
        });

    [UnmanagedCallersOnly(EntryPoint = "SQLGetDiagRecW")]
    internal static short SQLGetDiagRecW(short kind, nint id, short record, nint state, int* native, nint text, short capacity, short* length) =>
        Diagnostic(kind, id, record, state, native, text, capacity, length, true);

    [UnmanagedCallersOnly(EntryPoint = "SQLGetDiagRec")]
    internal static short SQLGetDiagRec(short kind, nint id, short record, nint state, int* native, nint text, short capacity, short* length) =>
        Diagnostic(kind, id, record, state, native, text, capacity, length, false);

    private static short Diagnostic(short kind, nint id, short record, nint state, int* native, nint text, short capacity, short* length, bool wide) {
        if (!DriverHandles.All.TryGetValue(id, out var handle) || handle.Kind != kind) return -2;
        lock (handle.Sync) {
            try {
                if (record < 1 || capacity < 0) return -1;
                if (record > handle.Diagnostics.Count) return 100;
                DriverDiagnostic diagnostic = handle.Diagnostics[record - 1];
                NativeText.Write(diagnostic.State, state, 6, null, wide);
                if (native != null) *native = 0;
                return NativeText.Write(diagnostic.Message, text, capacity, length, wide);
            } catch { return -1; }
        }
    }
}
