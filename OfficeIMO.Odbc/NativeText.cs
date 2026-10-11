using System.Runtime.InteropServices;
using System.Text;

namespace OfficeIMO.Odbc;

internal static unsafe class NativeText {
    private static readonly Encoding Ansi = CreateAnsiEncoding();
    private static Encoding CreateAnsiEncoding() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return OperatingSystem.IsWindows() ? Encoding.GetEncoding((int)GetACP()) : Encoding.UTF8;
    }

    [DllImport("kernel32.dll")] private static extern uint GetACP();

    internal static string Read(nint pointer, int length, bool wide) {
        if (pointer == 0) {
            if (length > 0) throw new ArgumentException("A non-empty string requires a buffer.");
            return "";
        }
        if (length < 0 && length != -3) throw new ArgumentException("Invalid native string length.");
        const int limit = 64 * 1024;
        if (length == -3) {
            length = 0;
            while (length <= limit && (wide ? ((ushort*)pointer)[length] != 0 : ((byte*)pointer)[length] != 0)) length++;
        }
        if (length > limit) throw new ArgumentException("Native string exceeds 64 Ki characters.");
        return wide ? new string((char*)pointer, 0, length) : Ansi.GetString(new ReadOnlySpan<byte>((void*)pointer, length));
    }

    /// <summary>Copies a terminated string; capacity and reported length are characters for wide ODBC APIs.</summary>
    internal static short Write(string value, nint buffer, int capacity, short* length, bool wide) {
        if (capacity < 0) throw new ArgumentException("Negative buffer length.");
        byte[] bytes = Bytes(value, wide);
        int units = wide ? bytes.Length / 2 : bytes.Length;
        if (length != null) *length = (short)Math.Min(short.MaxValue, units);
        int available = Math.Max(0, capacity - 1) * (wide ? 2 : 1);
        int copied = Math.Min(available, bytes.Length);
        if (buffer != 0 && capacity > 0) {
            bytes.AsSpan(0, copied).CopyTo(new Span<byte>((void*)buffer, copied));
            if (wide) ((ushort*)buffer)[copied / 2] = 0;
            else ((byte*)buffer)[copied] = 0;
        }
        return buffer != 0 && copied < bytes.Length ? (short)1 : (short)0;
    }

    internal static byte[] Bytes(string value, bool wide) => (wide ? Encoding.Unicode : Ansi).GetBytes(value);
}
