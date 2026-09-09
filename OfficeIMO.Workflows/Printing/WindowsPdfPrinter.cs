using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

[SupportedOSPlatform("windows")]
internal static partial class WindowsPdfPrinter {
    internal static Task<IReadOnlyList<PdfPrinterInfo>> GetPrintersAsync(CancellationToken token) => Task.Run(() => GetPrinters(token), token);
    internal static Task<PdfPrintSubmission> SubmitAsync(PdfPreparedPrintDocument document, PdfPrintDeliveryOptions options, CancellationToken token) =>
        Task.Run(() => Submit(document, options, token), CancellationToken.None);

    internal static IReadOnlyList<PdfPrinterInfo> GetPrinters(CancellationToken token) {
        token.ThrowIfCancellationRequested();
        uint needed = 0;
        _ = Native.EnumPrinters(6, null, 2, IntPtr.Zero, 0, out needed, out _);
        if (needed == 0) {
            int error = Marshal.GetLastWin32Error();
            if (error != 0 && error != 122) throw new Win32Exception(error);
            return [];
        }
        if (needed > 16 * 1024 * 1024) throw new InvalidOperationException("Printer discovery exceeded its memory limit.");
        IntPtr buffer = Marshal.AllocHGlobal(checked((int)needed));
        try {
            if (!Native.EnumPrinters(6, null, 2, buffer, needed, out _, out uint returned)) throw new Win32Exception();
            uint characters = 0;
            _ = Native.GetDefaultPrinter(null, ref characters);
            string? defaultName = null;
            if (characters is > 0 and < 32768) {
                var name = new System.Text.StringBuilder((int)characters);
                if (Native.GetDefaultPrinter(name, ref characters)) defaultName = name.ToString();
            }
            var printers = new List<PdfPrinterInfo>();
            int size = Marshal.SizeOf<Native.PrinterInfo2>();
            if (returned > needed / size) throw new InvalidOperationException("The printer list has invalid bounds.");
            for (int index = 0; index < returned; index++) {
                token.ThrowIfCancellationRequested();
                var info = Marshal.PtrToStructure<Native.PrinterInfo2>(IntPtr.Add(buffer, index * size));
                string name = Marshal.PtrToStringUni(info.PrinterName) ?? string.Empty;
                string port = Marshal.PtrToStringUni(info.PortName) ?? string.Empty;
                if (name.Length == 0) continue;
                bool file = port.Split(',').Any(value => value.Equals("PORTPROMPT:", StringComparison.OrdinalIgnoreCase) || value.Equals("FILE:", StringComparison.OrdinalIgnoreCase));
                printers.Add(new(name, string.Equals(name, defaultName, StringComparison.OrdinalIgnoreCase), file));
            }
            return printers.OrderByDescending(printer => printer.IsDefault).ThenBy(printer => printer.Name, StringComparer.OrdinalIgnoreCase).ToArray();
        } finally { Marshal.FreeHGlobal(buffer); }
    }

    internal static PdfPrintSubmission Submit(PdfPreparedPrintDocument document, PdfPrintDeliveryOptions options, CancellationToken token) {
        PdfPrinterInfo printer = GetPrinters(token).FirstOrDefault(item => string.Equals(item.Name, options.PrinterName, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("The selected printer is not installed.", nameof(options));
        string? output = ValidateOutput(document.SourcePath, printer, options.OutputFilePath);
        if (!Native.OpenPrinter(printer.Name, out IntPtr printerHandle, IntPtr.Zero)) throw new Win32Exception();
        IntPtr mode = IntPtr.Zero, dc = IntPtr.Zero;
        int jobId = 0;
        bool completed = false;
        Native.AbortProcedure abort = (_, _) => !token.IsCancellationRequested;
        try {
            int modeBytes = Native.DocumentProperties(IntPtr.Zero, printerHandle, printer.Name, IntPtr.Zero, IntPtr.Zero, 0);
            if (modeBytes < 102 || modeBytes > 1024 * 1024) throw new InvalidOperationException("The printer did not return valid settings.");
            mode = Marshal.AllocHGlobal(modeBytes);
            if (Native.DocumentProperties(IntPtr.Zero, printerHandle, printer.Name, mode, IntPtr.Zero, 2) != 1) throw new Win32Exception();
            ConfigureMode(mode, printerHandle, printer.Name, document.Sheets[0].Plan.PaperSize, options);
            bool duplex = Marshal.ReadInt16(mode, 94) is 2 or 3;
            dc = Native.CreateDC("WINSPOOL", printer.Name, null, mode);
            if (dc == IntPtr.Zero) throw new Win32Exception();
            ValidatePaper(dc, document.Sheets[0].Plan.PaperSize);
            if (Native.SetAbortProc(dc, abort) <= 0) throw new Win32Exception();
            token.ThrowIfCancellationRequested();
            var info = new Native.DocInfo { Size = Marshal.SizeOf<Native.DocInfo>(), DocumentName = options.DocumentName, Output = output };
            jobId = Native.StartDoc(dc, ref info);
            if (jobId <= 0) throw new Win32Exception();
            for (int copy = 0; copy < options.Copies; copy++) {
                foreach (PdfRenderedPrintSheet sheet in document.Sheets) {
                    token.ThrowIfCancellationRequested();
                    ConfigureMode(mode, printerHandle, printer.Name, sheet.Plan.PaperSize, options);
                    IntPtr reset = Native.ResetDC(dc, mode);
                    if (reset == IntPtr.Zero) throw new Win32Exception();
                    dc = reset;
                    if (Native.SetAbortProc(dc, abort) <= 0) throw new Win32Exception();
                    ValidatePaper(dc, sheet.Plan.PaperSize);
                    if (!OfficeRasterImageDecoder.TryDecode(sheet.Png, new OfficeRasterDecodeOptions {
                        MaximumDecodedPixels = 16_000_000, MaximumEncodedBytes = 128 * 1024 * 1024, CancellationToken = token
                    }, out OfficeRasterImage? raster, out _) || raster is null) throw new InvalidOperationException("A prepared sheet could not be decoded.");
                    byte[] pixels = raster.GetPixels();
                    for (int index = 0; index < pixels.Length; index += 4) (pixels[index], pixels[index + 2]) = (pixels[index + 2], pixels[index]);
                    var bitmap = new Native.BitmapInfo {
                        Size = 40, Width = raster.Width, Height = -raster.Height, Planes = 1, BitCount = 32,
                        ImageSize = checked((uint)pixels.Length)
                    };
                    token.ThrowIfCancellationRequested();
                    if (Native.StartPage(dc) <= 0) throw new Win32Exception();
                    int drawn = Native.StretchDIBits(dc, -Native.GetDeviceCaps(dc, 112), -Native.GetDeviceCaps(dc, 113),
                        Native.GetDeviceCaps(dc, 110), Native.GetDeviceCaps(dc, 111), 0, 0, raster.Width, raster.Height,
                        pixels, ref bitmap, 0, 0x00CC0020);
                    if (drawn == 0 || drawn == -1) throw new Win32Exception();
                    if (Native.EndPage(dc) <= 0) throw new Win32Exception();
                }
                // Start each collated copy on a fresh physical sheet when duplex leaves an unused back side.
                if (duplex && document.Sheets.Count % 2 != 0 && copy + 1 < options.Copies) {
                    token.ThrowIfCancellationRequested();
                    if (Native.StartPage(dc) <= 0 || Native.EndPage(dc) <= 0) throw new Win32Exception();
                }
            }
            token.ThrowIfCancellationRequested();
            if (Native.EndDoc(dc) <= 0) throw new Win32Exception();
            completed = true;
            return new(printer.Name, jobId.ToString(CultureInfo.InvariantCulture), document.Sheets.Count, options.Copies, output);
        } catch (Exception error) when (jobId > 0) {
            throw new PdfPrintDeliveryException(jobId.ToString(CultureInfo.InvariantCulture), error);
        } finally {
            if (dc != IntPtr.Zero) {
                if (jobId > 0 && !completed) _ = Native.AbortDoc(dc);
                _ = Native.DeleteDC(dc);
            }
            if (mode != IntPtr.Zero) Marshal.FreeHGlobal(mode);
            _ = Native.ClosePrinter(printerHandle);
            GC.KeepAlive(abort);
        }
    }

    private static string? ValidateOutput(string sourcePath, PdfPrinterInfo printer, string? outputPath) {
        if (!printer.RequiresOutputFile) {
            if (!string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("This printer does not accept an output-file path.");
            return null;
        }
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("Choose a new output file for this printer.");
        string output = OfficeIMO.Internal.OfficeStorageIdentity.GetLocalPath(outputPath)
            ?? throw new ArgumentException("A file printer requires a local output path.");
        output = Path.GetFullPath(output);
        if (OfficeIMO.Internal.OfficeStorageIdentity.AreEquivalent(output, sourcePath)) throw new IOException("A print job cannot replace its source document.");
        if (File.Exists(output) || Directory.Exists(output)) throw new IOException("The print output already exists. Choose a new filename.");
        if (!Directory.Exists(Path.GetDirectoryName(output))) throw new DirectoryNotFoundException("The print output directory does not exist.");
        return output;
    }

    private static void ConfigureMode(IntPtr mode, IntPtr printer, string name, PageSize paper, PdfPrintDeliveryOptions options) {
        // Public DEVMODEW offsets; preserve the driver's complete private trailing data.
        if ((ushort)Marshal.ReadInt16(mode, 68) < 102) throw new InvalidOperationException("Unsupported printer settings layout.");
        bool landscape = paper.Width > paper.Height;
        double width = Math.Min(paper.Width, paper.Height), height = Math.Max(paper.Width, paper.Height);
        short paperCode = PaperCode(width, height);
        int fields = Marshal.ReadInt32(mode, 72) | 1 | 2 | 0x10 | 0x100;
        Marshal.WriteInt16(mode, 76, landscape ? (short)2 : (short)1);
        Marshal.WriteInt16(mode, 78, paperCode);
        if (paperCode == 256) {
            fields |= 4 | 8;
            Marshal.WriteInt16(mode, 80, checked((short)Math.Round(height * 254 / 72)));
            Marshal.WriteInt16(mode, 82, checked((short)Math.Round(width * 254 / 72)));
        } else fields &= ~(4 | 8);
        Marshal.WriteInt16(mode, 84, 100);
        Marshal.WriteInt16(mode, 86, 1); // Software supplies collated copies exactly once.
        if (options.Duplex != PdfPrintDuplex.PrinterDefault) {
            fields |= 0x1000;
            Marshal.WriteInt16(mode, 94, (short)options.Duplex);
        }
        Marshal.WriteInt32(mode, 72, fields);
        if (Native.DocumentProperties(IntPtr.Zero, printer, name, mode, mode, 10) != 1) throw new Win32Exception();
        if (Marshal.ReadInt16(mode, 86) != 1 || (options.Duplex != PdfPrintDuplex.PrinterDefault && Marshal.ReadInt16(mode, 94) != (short)options.Duplex))
            throw new InvalidOperationException("The printer did not accept the requested copy or duplex settings.");
    }

    private static short PaperCode(double width, double height) {
        if (Math.Abs(width - PageSizes.A4.Width) < 1 && Math.Abs(height - PageSizes.A4.Height) < 1) return 9;
        if (Math.Abs(width - PageSizes.A3.Width) < 1 && Math.Abs(height - PageSizes.A3.Height) < 1) return 8;
        if (Math.Abs(width - PageSizes.Letter.Width) < 1 && Math.Abs(height - PageSizes.Letter.Height) < 1) return 1;
        if (Math.Abs(width - PageSizes.Legal.Width) < 1 && Math.Abs(height - PageSizes.Legal.Height) < 1) return 5;
        return 256;
    }

    private static void ValidatePaper(IntPtr dc, PageSize paper) {
        int dpiX = Native.GetDeviceCaps(dc, 88), dpiY = Native.GetDeviceCaps(dc, 90);
        if (dpiX <= 0 || dpiY <= 0 || Math.Abs(Native.GetDeviceCaps(dc, 110) * 72D / dpiX - paper.Width) > 3 ||
            Math.Abs(Native.GetDeviceCaps(dc, 111) * 72D / dpiY - paper.Height) > 3)
            throw new InvalidOperationException("The selected printer did not accept the reviewed paper size and orientation.");
    }
}
