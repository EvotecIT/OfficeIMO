using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace OfficeIMO.Workflows;

#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
internal static partial class WindowsPdfPrinter {
    // DeviceCapabilities indexes (wingdi.h).
    private const ushort DcOrientation = 17, DcPapers = 2, DcPaperSize = 3, DcBins = 6, DcDuplex = 7, DcBinNames = 12, DcEnumResolutions = 13,
        DcPaperNames = 16, DcCopies = 18, DcCollate = 22, DcColorDevice = 32;
    // Identifiers are WORDs. DeviceCapabilities has no output-capacity argument and the driver may change
    // its count between calls, so list buffers reserve the entire identifier space.
    private const int MaximumListEntries = 65536, MaximumResolutions = 4096;

    internal static Task<PdfPrinterCapabilities> GetCapabilitiesAsync(string printerName, PrinterCapabilityCache cache, CancellationToken token) =>
        Task.Run(() => GetCapabilities(printerName, cache, token), token);

    private static PdfPrinterCapabilities GetCapabilities(string printerName, PrinterCapabilityCache cache, CancellationToken token) {
        PrinterDetails printer = FindPrinter(printerName, token);
        PdfPrinterCapabilities capabilities = WindowsPrinterSettings.BuildCapabilities(printer.Info.Name, ReadDriverCapabilities(printer, token));
        cache.Set(printerName, capabilities);
        return capabilities;
    }

    private static PrinterDetails FindPrinter(string printerName, CancellationToken token) =>
        GetPrinterDetails(token).FirstOrDefault(printer => string.Equals(printer.Info.Name, printerName, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("The selected printer is not installed.", nameof(printerName));

    private static WindowsPrinterSettings.DriverCapabilities ReadDriverCapabilities(PrinterDetails printer, CancellationToken token) {
        string name = printer.Info.Name, port = printer.PortName;
        token.ThrowIfCancellationRequested();
        ushort[] codes = ReadWords(name, port, DcPapers);
        (int, int)[] sizes = ReadPairs(name, port, DcPaperSize, MaximumListEntries);
        string[] names = ReadStrings(name, port, DcPaperNames, 64);
        token.ThrowIfCancellationRequested();
        ushort[] bins = ReadWords(name, port, DcBins);
        string[] binNames = ReadStrings(name, port, DcBinNames, 24);
        // A driver can answer the bin list but not the names; those sources are then labelled by their identifiers.
        if (binNames.Length > 0 && bins.Length != binNames.Length) throw new InvalidOperationException("The printer's paper-source list changed during discovery. Refresh the printer list.");
        var sources = new List<PdfPaperSourceInfo>(bins.Length);
        for (int index = 0; index < bins.Length; index++) {
            string id = bins[index].ToString(CultureInfo.InvariantCulture);
            string label = binNames.Length > 0 ? binNames[index].Trim() : string.Empty;
            if (sources.All(source => source.Id != id)) sources.Add(new(id, label.Length == 0 ? id : label));
        }
        token.ThrowIfCancellationRequested();
        return new WindowsPrinterSettings.DriverCapabilities {
            PaperCodes = codes, PaperSizes = sizes, PaperNames = names, DefaultPaperCode = ReadDefaultPaperCode(name),
            PaperSources = sources,
            Duplex = Native.DeviceCapabilities(name, port, DcDuplex, IntPtr.Zero, IntPtr.Zero),
            ColorDevice = Native.DeviceCapabilities(name, port, DcColorDevice, IntPtr.Zero, IntPtr.Zero),
            Collate = Native.DeviceCapabilities(name, port, DcCollate, IntPtr.Zero, IntPtr.Zero),
            Copies = Native.DeviceCapabilities(name, port, DcCopies, IntPtr.Zero, IntPtr.Zero),
            Resolutions = ReadPairs(name, port, DcEnumResolutions, MaximumResolutions)
        };
    }

    private static int Count(string name, string port, ushort capability, int maximum) {
        int count = Native.DeviceCapabilities(name, port, capability, IntPtr.Zero, IntPtr.Zero);
        if (count < 0) return 0; // The driver does not report this list.
        if (count > maximum) throw new InvalidOperationException("The printer reported too many capability entries.");
        return count;
    }

    private static ushort[] ReadWords(string name, string port, ushort capability) {
        if (Count(name, port, capability, MaximumListEntries) == 0) return [];
        IntPtr buffer = Marshal.AllocHGlobal(MaximumListEntries * 2);
        try {
            int returned = Native.DeviceCapabilities(name, port, capability, buffer, IntPtr.Zero);
            if (returned < 0) return []; // The driver does not report this list.
            var values = new ushort[Math.Min(returned, MaximumListEntries)];
            for (int index = 0; index < values.Length; index++) values[index] = unchecked((ushort)Marshal.ReadInt16(buffer, index * 2));
            return values;
        } finally { Marshal.FreeHGlobal(buffer); }
    }

    private static (int, int)[] ReadPairs(string name, string port, ushort capability, int maximum) {
        int count = Count(name, port, capability, maximum);
        if (count == 0) return [];
        IntPtr buffer = Marshal.AllocHGlobal(maximum * 8);
        try {
            int returned = Native.DeviceCapabilities(name, port, capability, buffer, IntPtr.Zero);
            if (returned < 0) return []; // The driver does not report this list.
            if (returned > maximum) throw new InvalidOperationException("The printer reported too many capability entries.");
            var values = new (int, int)[returned];
            for (int index = 0; index < returned; index++) values[index] = (Marshal.ReadInt32(buffer, index * 8), Marshal.ReadInt32(buffer, index * 8 + 4));
            return values;
        } finally { Marshal.FreeHGlobal(buffer); }
    }

    private static string[] ReadStrings(string name, string port, ushort capability, int characters) {
        if (Count(name, port, capability, MaximumListEntries) == 0) return [];
        IntPtr buffer = Marshal.AllocHGlobal(MaximumListEntries * characters * 2);
        try {
            int returned = Native.DeviceCapabilities(name, port, capability, buffer, IntPtr.Zero);
            if (returned < 0) return []; // The driver does not report this list.
            var values = new string[Math.Min(returned, MaximumListEntries)];
            for (int index = 0; index < values.Length; index++)
                values[index] = (Marshal.PtrToStringUni(IntPtr.Add(buffer, index * characters * 2), characters) ?? string.Empty).Split('\0')[0];
            return values;
        } finally { Marshal.FreeHGlobal(buffer); }
    }

    private static short ReadDefaultPaperCode(string name) {
        if (!Native.OpenPrinter(name, out IntPtr handle, IntPtr.Zero)) throw new Win32Exception();
        IntPtr mode = IntPtr.Zero;
        try {
            int bytes = Native.DocumentProperties(IntPtr.Zero, handle, name, IntPtr.Zero, IntPtr.Zero, 0);
            if (bytes < 102 || bytes > 1024 * 1024) return 0;
            mode = Marshal.AllocHGlobal(bytes);
            if (Native.DocumentProperties(IntPtr.Zero, handle, name, mode, IntPtr.Zero, 2) != 1) return 0;
            return (Marshal.ReadInt32(mode, WindowsPrinterSettings.DevModeFields) & WindowsPrinterSettings.FieldPaperSize) != 0
                ? Marshal.ReadInt16(mode, WindowsPrinterSettings.DevModePaperSize) : (short)0;
        } finally {
            if (mode != IntPtr.Zero) Marshal.FreeHGlobal(mode);
            _ = Native.ClosePrinter(handle);
        }
    }
}
