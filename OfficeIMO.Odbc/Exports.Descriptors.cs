using System.Runtime.InteropServices;

namespace OfficeIMO.Odbc {

    internal static unsafe partial class Exports {
        [UnmanagedCallersOnly(EntryPoint = "SQLGetDescFieldW")]
        internal static short SQLGetDescFieldW(nint id, short record, short field, nint output, int capacity, int* length) =>
            GetDescriptorField(id, record, field, output, capacity, length, true);
        [UnmanagedCallersOnly(EntryPoint = "SQLGetDescField")]
        internal static short SQLGetDescField(nint id, short record, short field, nint output, int capacity, int* length) =>
            GetDescriptorField(id, record, field, output, capacity, length, false);

        private static short GetDescriptorField(nint id, short record, short field, nint output, int capacity, int* length, bool wide) =>
            DriverHandles.With<DescriptorHandle>(id, descriptor => {
                if (output == 0) return descriptor.Fail("HY009", "Output buffer is required.");
                StatementHandle statement = descriptor.Statement;
                if (field == 1099) { *(short*)output = 1; if (length != null) *length = 2; return 0; }
                if (field == 1001) {
                    int count = descriptor.Attribute switch {
                        10010 => statement.Bindings.Count == 0 ? 0 : statement.Bindings.Keys.Max(),
                        10011 => statement.Parameters.Count == 0 ? 0 : statement.Parameters.Keys.Max(),
                        10012 => statement.Reader?.FieldCount ?? 0,
                        _ => statement.ParameterCount
                    };
                    *(short*)output = (short)count;
                    if (length != null) *length = 2;
                    return 0;
                }
                if (field == 25) { *(int*)output = 0; if (length != null) *length = sizeof(int); return 0; }
                if (field is 20 or 21 or 24 or 34) {
                    *(nint*)output = field switch { 20 => 1, 21 when descriptor.Attribute == 10012 => statement.RowStatusPointer,
                        34 when descriptor.Attribute == 10012 => statement.RowsFetchedPointer, _ => 0 };
                    if (length != null) *length = sizeof(nint);
                    return 0;
                }
                if (descriptor.Attribute != 10012) return descriptor.Fail("HYC00", "Application descriptors are exposed for identity and scalar binding metadata only.");
                if (statement.Reader is not { } reader || record < 1 || record > reader.FieldCount)
                    return descriptor.Fail("07009", "Invalid descriptor record.");
                short type = NativeValue.SqlType(reader, record - 1);
                if (field == 1011) {
                    short actual = 0;
                    short result = NativeText.Write(reader.GetName(record - 1), output, wide ? capacity / 2 : capacity, &actual, wide);
                    if (length != null) *length = actual * (wide ? 2 : 1);
                    if (result == 1) descriptor.Diagnostics.Add(new("01004", "Descriptor name was truncated."));
                    return result;
                }
                if (field is 1002 or 1005 or 1006 or 1007 or 1008) {
                    *(short*)output = field switch { 1002 => type == 93 ? (short)9 : type, 1005 => NativeValue.Precision(type),
                        1006 => type == 93 ? (short)7 : (short)0, 1007 => type == 93 ? (short)3 : (short)0, _ => 1 };
                    if (length != null) *length = 2;
                    return 0;
                }
                if (field == 1003) { *(nuint*)output = NativeValue.ColumnSize(type); if (length != null) *length = sizeof(nuint); return 0; }
                return descriptor.Fail("HYC00", "Unsupported descriptor field.");
            });
    }
}
