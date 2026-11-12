using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;

namespace OfficeIMO.Workflows;

internal static partial class WindowsPdfPrinter {
    internal static Task<IReadOnlyList<PdfPaperSourceInfo>> GetPaperSourcesAsync(string printerName, CancellationToken token) =>
        Task.Run(() => GetPaperSources(printerName, token), token);

    private static IReadOnlyList<PdfPaperSourceInfo> GetPaperSources(string printerName, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var printer = GetPrinterDetails(token).FirstOrDefault(printer => string.Equals(printer.Info.Name, printerName, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("The selected printer is not installed.", nameof(printerName));
        // Bin identifiers are WORDs. Reserve the entire identifier space because DeviceCapabilities
        // has no output-capacity argument and the driver may change its count between calls.
        const int maximumBins = 65536;
        int count = Native.DeviceCapabilities(printerName, printer.PortName, 6, IntPtr.Zero, IntPtr.Zero);
        if (count < 0) throw new Win32Exception();
        if (count == 0) return [];
        if (count > maximumBins) throw new InvalidOperationException("The printer reported too many paper sources.");
        IntPtr bins = Marshal.AllocHGlobal(maximumBins * 2);
        IntPtr names = IntPtr.Zero;
        try {
            names = Marshal.AllocHGlobal(maximumBins * 24 * 2);
            int returned = Native.DeviceCapabilities(printerName, printer.PortName, 6, bins, IntPtr.Zero);
            int named = Native.DeviceCapabilities(printerName, printer.PortName, 12, names, IntPtr.Zero);
            if (returned < 0 || named < 0) throw new Win32Exception();
            if (returned > maximumBins || returned != named)
                throw new InvalidOperationException("The printer's paper-source list changed during discovery. Refresh the printer list.");
            var sources = new List<PdfPaperSourceInfo>();
            for (int index = 0; index < returned; index++) {
                token.ThrowIfCancellationRequested();
                ushort bin = unchecked((ushort)Marshal.ReadInt16(bins, index * 2));
                string name = (Marshal.PtrToStringUni(IntPtr.Add(names, index * 48), 24) ?? string.Empty).Split('\0')[0].Trim();
                string id = bin.ToString(CultureInfo.InvariantCulture);
                sources.Add(new(id, name.Length == 0 ? id : name));
            }
            return sources.DistinctBy(source => source.Id).ToArray();
        } finally {
            if (names != IntPtr.Zero) Marshal.FreeHGlobal(names);
            Marshal.FreeHGlobal(bins);
        }
    }
}
