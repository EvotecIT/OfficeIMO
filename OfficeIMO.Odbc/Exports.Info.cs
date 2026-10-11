using System.Runtime.InteropServices;

namespace OfficeIMO.Odbc;

internal static unsafe partial class Exports {
    // ODBC function identifiers are defined in sql.h and sqlext.h, independently of ANSI/W entrypoints.
    private static readonly ushort[] SupportedFunctions = [4, 5, 6, 8, 9, 11, 12, 13, 16, 18, 19, 20, 40, 41, 43, 44, 45, 47, 54, 61, 63, 72,
        1001, 1003, 1005, 1006, 1007, 1008, 1010, 1011, 1012, 1014, 1016, 1019, 1020, 1021];

    [UnmanagedCallersOnly(EntryPoint = "SQLGetFunctions")]
    internal static short SQLGetFunctions(nint id, ushort function, ushort* supported) => DriverHandles.With<ConnectionHandle>(id, handle => {
        if (supported == null) return handle.Fail("HY009", "Output buffer is required.");
        if (function == 999) {
            new Span<ushort>(supported, 250).Clear();
            foreach (ushort item in SupportedFunctions) supported[item >> 4] |= (ushort)(1 << (item & 15));
        } else if (function == 0) {
            new Span<ushort>(supported, 100).Clear();
            foreach (ushort item in SupportedFunctions.Where(item => item < 100)) supported[item] = 1;
        } else *supported = SupportedFunctions.Contains(function) ? (ushort)1 : (ushort)0;
        return 0;
    });

    [UnmanagedCallersOnly(EntryPoint = "SQLGetInfoW")]
    internal static short SQLGetInfoW(nint id, ushort info, nint output, short capacity, short* length) => GetInfo(id, info, output, capacity, length, true);
    [UnmanagedCallersOnly(EntryPoint = "SQLGetInfo")]
    internal static short SQLGetInfo(nint id, ushort info, nint output, short capacity, short* length) => GetInfo(id, info, output, capacity, length, false);

    private static short GetInfo(nint id, ushort info, nint output, short capacity, short* length, bool wide) =>
        DriverHandles.With<ConnectionHandle>(id, handle => {
            string? text = info switch {
                2 => "", 6 => "OfficeIMO.Odbc.dll", 7 => "03.04.0005", 10 or 77 => "03.00",
                13 => "OfficeIMO", 14 => "\\", 16 => handle.Snapshot is null ? "" : System.IO.Path.GetFileName(handle.Snapshot.Path),
                17 => "OfficeIMO Excel Snapshot", 18 => "01.00.0000", 19 or 25 or 27 or 38 or 87 => "Y",
                11 or 20 or 21 or 36 or 37 or 73 or 90 or 103 or 111 or 113 or 10002 or 10003 => "N",
                29 => "\"", 39 or 40 or 42 or 47 or 89 or 94 => "", 41 => ".", 45 => "table",
                10000 => "1995", 10004 => "BINARY", _ => null
            };
            if (text is not null) {
                short units = 0;
                short result = NativeText.Write(text, output, wide ? capacity / 2 : capacity, &units, wide);
                if (length != null) *length = checked((short)(units * (wide ? 2 : 1)));
                if (result == 1) handle.Diagnostics.Add(new("01004", "Driver information was truncated."));
                return result;
            }
            if (output == 0) return handle.Fail("HY009", "Output buffer is required.");
            // SQLUSMALLINT-valued information types must not overwrite a two-byte caller buffer.
            if (info is 0 or 1 or 9 or 12 or 15 or 22 or 23 or 24 or 28 or 30 or 31 or 32 or 33 or 34 or 35 or 46 or 74 or 75 or 84 or 85 or 88 or 93 or 97 or 98 or 99 or 100 or 101 or 106 or 107 or 114 or 10005) {
                *(ushort*)output = info switch { 0 => 32, 1 => 0, 28 => 4, 93 => 3, 23 or 24 => 2,
                    30 or 35 or 10005 => 255, 74 => 2, 75 => 1, 84 => 1, 85 => 1, 88 => 3,
                    97 or 99 or 100 or 101 => 256, 106 => 64, 114 => 1, _ => 0 };
                if (length != null) *length = 2;
                return 0;
            }
            uint? number = info switch {
                26 or 72 or 78 or 79 or 80 or 82 or 83 or 86 or 91 or 92 or 118 or 119 or 120 or 121 or 10001 or 10021 => 0,
                48 or 49 or 50 or 51 or 52 => 0,
                8 or 43 or 44 => 1, // Forward-only, read-only cursors.
                81 => 11, // SQL_GD_ANY_COLUMN | SQL_GD_ANY_ORDER | SQL_GD_BOUND.
                95 => 31, 96 => 3, 115 => 127,
                102 or 104 => 0,
                105 => 64 * 1024,
                152 => 0,
                _ => null
            };
            if (number is null) return handle.Fail("HY096", "Unsupported information type.");
            *(uint*)output = number.Value;
            if (length != null) *length = 4;
            return 0;
        });
}
