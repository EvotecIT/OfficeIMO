using System.ComponentModel;
using System.Globalization;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>Platform-independent DEVMODEW composition and capability interpretation, kept apart from GDI calls so it can be verified anywhere.</summary>
internal static partial class WindowsPrinterSettings {
    /// <summary>
    /// The device page for each step and how far its sheet turns onto it. With <paramref name="turnSheets"/>, a sheet on the same paper
    /// turned the other way stays on the current device page, rotated as the driver turns landscape (<paramref name="landscapeTurn"/>
    /// degrees counter-clockwise, 90 or 270), so it lands upright on the fed paper. A different paper size starts a new device page.
    /// </summary>
    internal static (PdfPageSize Device, int QuarterTurns)[] DeviceLayout(IReadOnlyList<PdfPrintSheet?> sheets, bool turnSheets, int landscapeTurn) {
        var layout = new (PdfPageSize, int)[sheets.Count];
        PdfPageSize? device = null;
        int turn = landscapeTurn == 270 ? 3 : 1;
        for (int index = 0; index < sheets.Count; index++) {
            if (sheets[index] is not PdfPrintSheet sheet) {
                layout[index] = (device ?? sheets.First(item => item is not null)!.PaperSize, 0);
                continue;
            }
            if (device is PdfPageSize current && SamePaper(sheet.PaperSize, current)) {
                layout[index] = (current, 0);
            } else if (turnSheets && device is PdfPageSize page && SameTurnedPaper(sheet.PaperSize, page)) {
                // A landscape sheet on a portrait page turns as the driver turns landscape; a portrait sheet on a landscape page turns back.
                bool landscapeSheet = sheet.PaperSize.Width.ToPoints() > sheet.PaperSize.Height.ToPoints();
                layout[index] = (page, landscapeSheet ? turn : 4 - turn);
            } else {
                device = sheet.PaperSize;
                layout[index] = (sheet.PaperSize, 0);
            }
        }
        return layout;
    }

    internal static bool SamePaper(PdfPageSize first, PdfPageSize second) =>
        Math.Abs(first.Width.ToPoints() - second.Width.ToPoints()) < 0.01 && Math.Abs(first.Height.ToPoints() - second.Height.ToPoints()) < 0.01;

    /// <summary>Whether <paramref name="first"/> is <paramref name="second"/> turned a quarter, such as A4 landscape and A4 portrait.</summary>
    internal static bool SameTurnedPaper(PdfPageSize first, PdfPageSize second) =>
        Math.Abs(first.Width.ToPoints() - second.Height.ToPoints()) < 0.01 && Math.Abs(first.Height.ToPoints() - second.Width.ToPoints()) < 0.01;

    // Public DEVMODEW offsets and dmFields flags (wingdi.h).
    internal const int DevModeSize = 68, DevModeFields = 72, DevModeOrientation = 76, DevModePaperSize = 78, DevModePaperLength = 80,
        DevModePaperWidth = 82, DevModeScale = 84, DevModeCopies = 86, DevModeDefaultSource = 88, DevModeDuplex = 94, DevModeCollate = 100;
    internal const int FieldOrientation = 0x1, FieldPaperSize = 0x2, FieldPaperLength = 0x4, FieldPaperWidth = 0x8, FieldScale = 0x10,
        FieldCopies = 0x100, FieldDefaultSource = 0x200, FieldDuplex = 0x1000, FieldCollate = 0x8000;
    internal const short CustomPaperCode = 256;

    /// <summary>Settings for one run of sheets with the same paper.</summary>
    internal readonly struct DevModeRequest {
        internal DevModeRequest(PdfPageSize paper, short paperCode, PdfPrintDuplex duplex, string? paperSourceId, short copies) {
            Paper = paper; PaperCode = paperCode; Duplex = duplex; PaperSourceId = paperSourceId; Copies = copies;
        }
        internal PdfPageSize Paper { get; }
        internal short PaperCode { get; }
        internal PdfPrintDuplex Duplex { get; }
        internal string? PaperSourceId { get; }
        /// <summary>Driver-produced collated copies; 1 when software repeats the sheets.</summary>
        internal short Copies { get; }
        internal bool Landscape => Paper.Width.ToPoints() > Paper.Height.ToPoints();
    }

    /// <summary>Writes the request into a driver-initialized DEVMODEW, preserving the driver's private trailing data.</summary>
    internal static void ComposeDevMode(byte[] mode, DevModeRequest request) {
        if (mode.Length < 102 || ReadUInt16(mode, DevModeSize) < 102) throw new InvalidOperationException("Unsupported printer settings layout.");
        int fields = ReadInt32(mode, DevModeFields) | FieldOrientation | FieldPaperSize | FieldScale | FieldCopies;
        WriteInt16(mode, DevModeOrientation, request.Landscape ? (short)2 : (short)1);
        WriteInt16(mode, DevModePaperSize, request.PaperCode);
        if (request.PaperCode == CustomPaperCode) {
            double width = Math.Min(request.Paper.Width.ToPoints(), request.Paper.Height.ToPoints());
            double height = Math.Max(request.Paper.Width.ToPoints(), request.Paper.Height.ToPoints());
            fields |= FieldPaperLength | FieldPaperWidth;
            WriteInt16(mode, DevModePaperLength, checked((short)Math.Round(height * 254 / 72)));
            WriteInt16(mode, DevModePaperWidth, checked((short)Math.Round(width * 254 / 72)));
        } else fields &= ~(FieldPaperLength | FieldPaperWidth);
        WriteInt16(mode, DevModeScale, 100);
        WriteInt16(mode, DevModeCopies, request.Copies);
        if (request.Copies > 1) {
            fields |= FieldCollate;
            WriteInt16(mode, DevModeCollate, 1);
        }
        if (request.PaperSourceId is not null) {
            fields |= FieldDefaultSource;
            WriteInt16(mode, DevModeDefaultSource, unchecked((short)ushort.Parse(request.PaperSourceId, CultureInfo.InvariantCulture)));
        }
        if (request.Duplex != PdfPrintDuplex.PrinterDefault) {
            fields |= FieldDuplex;
            WriteInt16(mode, DevModeDuplex, (short)request.Duplex);
        }
        WriteInt32(mode, DevModeFields, fields);
    }

    /// <summary>Describes the first setting the driver changed while merging the request, or null when it accepted all of them.</summary>
    /// <remarks>A driver that clears a requested setting's dmFields flag ignores the value, so that setting counts as rejected.</remarks>
    internal static string? FindRejectedSetting(byte[] mode, DevModeRequest request) {
        int fields = ReadInt32(mode, DevModeFields);
        if (ReadInt16(mode, DevModeCopies) != request.Copies || (request.Copies > 1 && (fields & FieldCopies) == 0)) return "copies";
        if (request.Copies > 1 && ((fields & FieldCollate) == 0 || ReadInt16(mode, DevModeCollate) != 1)) return "collation";
        if (request.Duplex != PdfPrintDuplex.PrinterDefault &&
            ((fields & FieldDuplex) == 0 || ReadInt16(mode, DevModeDuplex) != (short)request.Duplex)) return "duplex";
        if (request.PaperSourceId is not null && ((fields & FieldDefaultSource) == 0 ||
            unchecked((ushort)ReadInt16(mode, DevModeDefaultSource)).ToString(CultureInfo.InvariantCulture) != request.PaperSourceId)) return "paper source";
        return null;
    }

    /// <summary>Uses the driver's own code for the paper size, then the standard DMPAPER codes, then a custom size.</summary>
    internal static short ResolvePaperCode(PdfPageSize paper, IReadOnlyList<PdfPrinterMedia> driverPapers) {
        PdfPrinterMedia? match = PdfPrinterMediaNames.Match(driverPapers, null, paper);
        if (match is not null && short.TryParse(match.Id, NumberStyles.None, CultureInfo.InvariantCulture, out short code) && code > 0 && code != CustomPaperCode) return code;
        foreach ((short standard, double width, double height) in StandardPapers)
            if (PdfPrinterMediaNames.SameSize(paper, PdfPrinterMediaNames.Portrait(width, height))) return standard;
        return CustomPaperCode;
    }

    // DMPAPER codes with portrait sizes in points.
    private static readonly (short Code, double Width, double Height)[] StandardPapers = [
        (1, 612, 792), (5, 612, 1008), (3, 792, 1224), (7, 522, 756), (6, 396, 612),
        (8, 841.89, 1190.55), (9, 595.28, 841.89), (11, 419.53, 595.28), (70, 297.64, 419.53),
        (12, 728.5, 1031.81), (13, 515.91, 728.5), (20, 296.79, 684), (27, 311.81, 623.62), (28, 459.21, 649.13)
    ];

    /// <summary>
    /// Describes a failed GDI print call. StartPage, EndPage and StretchDIBits usually leave the last error at zero, so the Win32
    /// error is reported only when the call set one.
    /// </summary>
    internal static Exception GdiFailure(string call, int lastError) => lastError == 0
        ? new InvalidOperationException($"The printer driver call {call} failed.")
        : new Win32Exception(lastError, $"The printer driver call {call} failed: {new Win32Exception(lastError).Message}");

    private static ushort ReadUInt16(byte[] buffer, int offset) => (ushort)(buffer[offset] | buffer[offset + 1] << 8);
    private static short ReadInt16(byte[] buffer, int offset) => unchecked((short)ReadUInt16(buffer, offset));
    private static int ReadInt32(byte[] buffer, int offset) => buffer[offset] | buffer[offset + 1] << 8 | buffer[offset + 2] << 16 | buffer[offset + 3] << 24;
    private static void WriteInt16(byte[] buffer, int offset, short value) {
        buffer[offset] = unchecked((byte)value); buffer[offset + 1] = unchecked((byte)(value >> 8));
    }
    private static void WriteInt32(byte[] buffer, int offset, int value) {
        for (int index = 0; index < 4; index++) buffer[offset + index] = unchecked((byte)(value >> (8 * index)));
    }
}
