using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

#if NET8_0_OR_GREATER
[SupportedOSPlatform("windows")]
#endif
internal static partial class WindowsPdfPrinter {
    internal static Task<IReadOnlyList<PdfPrinterInfo>> GetPrintersAsync(CancellationToken token) => Task.Run(() => GetPrinters(token), token);
    internal static Task<PdfPrintSubmission> SubmitAsync(PdfPrintJob job, PdfPrintDeliveryOptions options, CancellationToken token) =>
        Task.Run(() => Submit(job, options, token), CancellationToken.None);

    internal static IReadOnlyList<PdfPrinterInfo> GetPrinters(CancellationToken token) => GetPrinterDetails(token).Select(printer => printer.Info).ToArray();

    private sealed record PrinterDetails(PdfPrinterInfo Info, string PortName);

    private static IReadOnlyList<PrinterDetails> GetPrinterDetails(CancellationToken token) {
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
            var printers = new List<PrinterDetails>();
            int size = Marshal.SizeOf<Native.PrinterInfo2>();
            if (returned > needed / size) throw new InvalidOperationException("The printer list has invalid bounds.");
            for (int index = 0; index < returned; index++) {
                token.ThrowIfCancellationRequested();
                var info = Marshal.PtrToStructure<Native.PrinterInfo2>(IntPtr.Add(buffer, index * size));
                string name = Marshal.PtrToStringUni(info.PrinterName) ?? string.Empty;
                string port = Marshal.PtrToStringUni(info.PortName) ?? string.Empty;
                if (name.Length == 0) continue;
                bool file = port.Split(',').Any(value => value.Equals("PORTPROMPT:", StringComparison.OrdinalIgnoreCase) || value.Equals("FILE:", StringComparison.OrdinalIgnoreCase));
                printers.Add(new(new(name, string.Equals(name, defaultName, StringComparison.OrdinalIgnoreCase), file), port));
            }
            return printers.OrderByDescending(printer => printer.Info.IsDefault).ThenBy(printer => printer.Info.Name, StringComparer.OrdinalIgnoreCase).ToArray();
        } finally { Marshal.FreeHGlobal(buffer); }
    }

    internal static PdfPrintSubmission Submit(PdfPrintJob job, PdfPrintDeliveryOptions options, CancellationToken token) {
        PrinterDetails details = FindPrinter(options.PrinterName, token);
        PdfPrinterInfo printer = details.Info;
        string? output = ValidateOutput(job.SourcePath, printer, options.OutputFilePath);
        PdfPrinterCapabilities capabilities = WindowsPrinterSettings.BuildCapabilities(printer.Name, ReadDriverCapabilities(details, token));
        if (options.PaperSourceId is not null && !capabilities.PaperSources.Any(source => source.Id == options.PaperSourceId))
            throw new ArgumentException("The selected paper source is no longer available on this printer.", nameof(options));
        IReadOnlyList<PdfPrintSheet> sheets = job.Plan.Sheets;
        bool singlePaper = sheets.All(sheet => SamePaper(sheet.PaperSize, sheets[0].PaperSize));
        short driverCopies = WindowsPrinterSettings.DriverCopies(options.Copies, capabilities, singlePaper);
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
            var driverDefaults = new byte[modeBytes];
            Marshal.Copy(mode, driverDefaults, 0, modeBytes);
            WindowsPrinterSettings.DevModeRequest Request(PdfPageSize paper, short copies) =>
                new(paper, WindowsPrinterSettings.ResolvePaperCode(paper, capabilities.Media), options.Duplex, options.PaperSourceId, copies);
            if (driverCopies > 1 && !TryConfigureMode(mode, driverDefaults, printerHandle, printer.Name, Request(sheets[0].PaperSize, driverCopies), out _))
                driverCopies = 1; // The driver declined to produce the copies; repeat the sheets instead.
            ConfigureMode(mode, driverDefaults, printerHandle, printer.Name, Request(sheets[0].PaperSize, driverCopies));
            PdfPageSize currentPaper = sheets[0].PaperSize;
            bool duplex = Marshal.ReadInt16(mode, WindowsPrinterSettings.DevModeDuplex) is 2 or 3;
            dc = Native.CreateDC("WINSPOOL", printer.Name, null, mode);
            if (dc == IntPtr.Zero) throw GdiFailure("CreateDC");
            ValidatePaper(dc, currentPaper);
            if (Native.SetAbortProc(dc, abort) <= 0) throw GdiFailure("SetAbortProc");
            token.ThrowIfCancellationRequested();
            var info = new Native.DocInfo { Size = Marshal.SizeOf<Native.DocInfo>(), DocumentName = options.DocumentName, Output = output };
            jobId = Native.StartDoc(dc, ref info);
            if (jobId <= 0) throw GdiFailure("StartDoc");
            int softwareCopies = options.Copies / driverCopies;
            var kept = new PrintSheetCache<PackedSheetPixels>(softwareCopies > 1 ? job.MaximumContentBytes : 0);
            // A null sheet is the blank back side that starts each collated copy on a fresh physical sheet under duplex.
            var steps = new List<(PdfPrintSheet? Sheet, bool First)>();
            for (int copy = 0; copy < softwareCopies; copy++) {
                foreach (PdfPrintSheet sheet in sheets) steps.Add((sheet, copy == 0));
                if (duplex && sheets.Count % 2 != 0 && copy + 1 < softwareCopies) steps.Add((null, false));
            }
            bool keep = softwareCopies > 1;
            bool gray = job.Plan.ColorMode == PdfPrintColorMode.Grayscale;
            // Sheets render a few ahead on other cores at the device resolution read now; a paper change that alters it renders again below.
            int deviceDpi = DeviceDpi(dc);
            double firstDpi = SheetDpi(deviceDpi, job, sheets[0], options.MaximumRasterDpi);
            long sheetBytes = (long)PdfPrintRaster.GdiStride((int)Math.Ceiling(sheets[0].PaperSize.Width.ToPoints() / 72 * firstDpi), gray) *
                (long)Math.Ceiling(sheets[0].PaperSize.Height.ToPoints() / 72 * firstDpi);
            using var ahead = new PrintSheetPrefetcher<SheetPixels>(steps.Count, PrintSheetPrefetcher<SheetPixels>.DepthFor(sheetBytes, job.MaximumContentBytes),
                (index, stop) => steps[index] is { Sheet: PdfPrintSheet sheet, First: true }
                    ? PrepareSheet(job, sheet, SheetDpi(deviceDpi, job, sheet, options.MaximumRasterDpi), gray, kept, keep, stop)
                    : null, token);
            for (int index = 0; index < steps.Count; index++) {
                token.ThrowIfCancellationRequested();
                SheetPixels? prepared = ahead.Take(index);
                PdfPrintSheet? sheet = steps[index].Sheet;
                if (sheet is null) {
                    if (Native.StartPage(dc) <= 0) throw GdiFailure("StartPage");
                    if (Native.EndPage(dc) <= 0) throw GdiFailure("EndPage");
                    continue;
                }
                if (!SamePaper(sheet.PaperSize, currentPaper)) {
                    ConfigureMode(mode, driverDefaults, printerHandle, printer.Name, Request(sheet.PaperSize, driverCopies));
                    IntPtr reset = Native.ResetDC(dc, mode);
                    if (reset == IntPtr.Zero) throw GdiFailure("ResetDC");
                    dc = reset;
                    if (Native.SetAbortProc(dc, abort) <= 0) throw GdiFailure("SetAbortProc");
                    ValidatePaper(dc, sheet.PaperSize);
                    currentPaper = sheet.PaperSize;
                }
                double dpi = SheetDpi(DeviceDpi(dc), job, sheet, options.MaximumRasterDpi);
                SheetPixels pixels = prepared is not null && prepared.Dpi == dpi ? prepared : PrepareSheet(job, sheet, dpi, gray, kept, keep, token);
                prepared = null;
                if (Native.StartPage(dc) <= 0) throw GdiFailure("StartPage");
                DrawPixels(dc, pixels, gray);
                if (Native.EndPage(dc) <= 0) throw GdiFailure("EndPage");
            }
            token.ThrowIfCancellationRequested();
            if (Native.EndDoc(dc) <= 0) throw GdiFailure("EndDoc");
            completed = true;
            return new(printer.Name, jobId.ToString(CultureInfo.InvariantCulture), sheets.Count, options.Copies, output);
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

    private sealed record SheetPixels(byte[] Rows, int Width, int Height, double Dpi);

    private static int DeviceDpi(IntPtr dc) => Math.Min(Native.GetDeviceCaps(dc, LogPixelsX), Native.GetDeviceCaps(dc, LogPixelsY));

    /// <summary>The sheet's raster resolution: the device resolution within the caller and pixel limits.</summary>
    private static double SheetDpi(int deviceDpi, PdfPrintJob job, PdfPrintSheet sheet, double maximumDpi) =>
        PdfPrintRaster.ResolveDpi(deviceDpi, maximumDpi, sheet.PaperSize.Width.ToPoints(), sheet.PaperSize.Height.ToPoints(), job.MaximumPixelsPerImage);

    /// <summary>
    /// Rasterizes one sheet to GDI bitmap rows without touching the device context, so it can run on another core while GDI spools
    /// earlier sheets. A sheet that a later software copy prints again is kept packed in <paramref name="kept"/>, so it renders once per job.
    /// </summary>
    private static SheetPixels PrepareSheet(PdfPrintJob job, PdfPrintSheet sheet, double dpi, bool gray, PrintSheetCache<PackedSheetPixels> kept,
        bool keep, CancellationToken token) {
        if (kept.TryGet(sheet.SheetNumber, out PackedSheetPixels packed) && packed.Dpi == dpi) return new(packed.Unpack(), packed.Width, packed.Height, dpi);
        OfficeRasterImage raster = job.RenderRaster(sheet.SheetNumber, dpi, job.MaximumPixelsPerImage, token).Raster;
        byte[] rows = PdfPrintRaster.ToGdiRows(raster.PixelBuffer, raster.Width, raster.Height, gray, token);
        if (keep) {
            PackedSheetPixels packing = PackedSheetPixels.Pack(rows, raster.Width, raster.Height, dpi);
            kept.Add(sheet.SheetNumber, packing, packing.PackedLength);
        }
        return new(rows, raster.Width, raster.Height, dpi);
    }

    /// <summary>Stretches the sheet's rows, 24-bit color or 8-bit gray, over the whole physical page of the started page.</summary>
    private static void DrawPixels(IntPtr dc, SheetPixels sheet, bool gray) {
        var header = new Native.BitmapInfo {
            Size = 40, Width = sheet.Width, Height = -sheet.Height, Planes = 1, BitCount = (ushort)(gray ? 8 : 24),
            ImageSize = checked((uint)sheet.Rows.Length), ColorsUsed = gray ? 256u : 0u
        };
        int x = -Native.GetDeviceCaps(dc, PhysicalOffsetX), y = -Native.GetDeviceCaps(dc, PhysicalOffsetY);
        int width = Native.GetDeviceCaps(dc, PhysicalWidth), height = Native.GetDeviceCaps(dc, PhysicalHeight);
        int drawn;
        if (gray) {
            var info = new Native.PaletteBitmapInfo { Header = header, Colors = GrayPalette };
            drawn = Native.StretchDIBits(dc, x, y, width, height, 0, 0, sheet.Width, sheet.Height, sheet.Rows, ref info, 0, 0x00CC0020);
        } else {
            drawn = Native.StretchDIBits(dc, x, y, width, height, 0, 0, sheet.Width, sheet.Height, sheet.Rows, ref header, 0, 0x00CC0020);
        }
        if (drawn == 0 || drawn == -1) throw GdiFailure("StretchDIBits");
    }

    private static readonly uint[] GrayPalette = Enumerable.Range(0, 256).Select(level => (uint)(level | level << 8 | level << 16)).ToArray();

    /// <summary>Names the failed GDI call; most print calls do not set the last error, which would otherwise read as success.</summary>
    private static Exception GdiFailure(string call) => WindowsPrinterSettings.GdiFailure(call, Marshal.GetLastWin32Error());

    private const int LogPixelsX = 88, LogPixelsY = 90, PhysicalWidth = 110, PhysicalHeight = 111, PhysicalOffsetX = 112, PhysicalOffsetY = 113;

    private static bool SamePaper(PdfPageSize first, PdfPageSize second) =>
        Math.Abs(first.Width.ToPoints() - second.Width.ToPoints()) < 0.01 && Math.Abs(first.Height.ToPoints() - second.Height.ToPoints()) < 0.01;

    private static string? ValidateOutput(string sourcePath, PdfPrinterInfo printer, string? outputPath) {
        if (!printer.RequiresOutputFile) {
            if (!(outputPath == null || string.IsNullOrWhiteSpace(outputPath))) throw new ArgumentException("This printer does not accept an output-file path.");
            return null;
        }
        if ((outputPath == null || string.IsNullOrWhiteSpace(outputPath))) throw new ArgumentException("Choose a new output file for this printer.");
        string output = OfficeIMO.Internal.OfficeStorageIdentity.GetLocalPath(outputPath)
            ?? throw new ArgumentException("A file printer requires a local output path.");
        output = Path.GetFullPath(output);
        if (OfficeIMO.Internal.OfficeStorageIdentity.AreEquivalent(output, sourcePath)) throw new IOException("A print job cannot replace its source document.");
        if (File.Exists(output) || Directory.Exists(output)) throw new IOException("The print output already exists. Choose a new filename.");
        if (!Directory.Exists(Path.GetDirectoryName(output))) throw new DirectoryNotFoundException("The print output directory does not exist.");
        return output;
    }

    private static void ConfigureMode(IntPtr mode, byte[] driverDefaults, IntPtr printer, string name, WindowsPrinterSettings.DevModeRequest request) {
        if (!TryConfigureMode(mode, driverDefaults, printer, name, request, out string? rejected))
            throw new InvalidOperationException("The printer did not accept the requested " + rejected + " setting.");
    }

    /// <summary>Starts from the driver defaults, applies the request, lets the driver merge it, and checks the merged result.</summary>
    private static bool TryConfigureMode(IntPtr mode, byte[] driverDefaults, IntPtr printer, string name,
        WindowsPrinterSettings.DevModeRequest request, out string? rejected) {
        byte[] composed = (byte[])driverDefaults.Clone();
        WindowsPrinterSettings.ComposeDevMode(composed, request);
        Marshal.Copy(composed, 0, mode, composed.Length);
        if (Native.DocumentProperties(IntPtr.Zero, printer, name, mode, mode, 10) != 1) throw new Win32Exception();
        Marshal.Copy(mode, composed, 0, composed.Length);
        rejected = WindowsPrinterSettings.FindRejectedSetting(composed, request);
        return rejected is null;
    }

    private static void ValidatePaper(IntPtr dc, PdfPageSize paper) {
        int dpiX = Native.GetDeviceCaps(dc, LogPixelsX), dpiY = Native.GetDeviceCaps(dc, LogPixelsY);
        if (dpiX <= 0 || dpiY <= 0 || Math.Abs(Native.GetDeviceCaps(dc, PhysicalWidth) * 72D / dpiX - paper.Width.ToPoints()) > 3 ||
            Math.Abs(Native.GetDeviceCaps(dc, PhysicalHeight) * 72D / dpiY - paper.Height.ToPoints()) > 3)
            throw new InvalidOperationException("The selected printer did not accept the reviewed paper size and orientation.");
    }
}
