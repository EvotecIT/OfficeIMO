using System.Globalization;

namespace OfficeIMO.Workflows;

internal static partial class WindowsPrinterSettings {
    /// <summary>Raw <c>DeviceCapabilities</c> results for one queue.</summary>
    internal sealed class DriverCapabilities {
        internal IReadOnlyList<ushort> PaperCodes { get; init; } = [];
        /// <summary>DC_PAPERSIZE entries in tenths of a millimetre.</summary>
        internal IReadOnlyList<(int Width, int Height)> PaperSizes { get; init; } = [];
        internal IReadOnlyList<string> PaperNames { get; init; } = [];
        /// <summary>dmPaperSize of the driver's default DEVMODE.</summary>
        internal short DefaultPaperCode { get; init; }
        internal IReadOnlyList<PdfPaperSourceInfo> PaperSources { get; init; } = [];
        internal int Duplex { get; init; } = -1;
        internal int ColorDevice { get; init; } = -1;
        internal int Collate { get; init; } = -1;
        internal int Copies { get; init; } = -1;
        internal IReadOnlyList<(int X, int Y)> Resolutions { get; init; } = [];
    }

    /// <summary>Interprets DeviceCapabilities results. Negative results mean the driver did not answer.</summary>
    internal static PdfPrinterCapabilities BuildCapabilities(string printerName, DriverCapabilities driver) {
        if (driver.PaperSizes.Count != driver.PaperCodes.Count || driver.PaperNames.Count != driver.PaperCodes.Count)
            throw new InvalidOperationException("The printer's paper list changed during discovery. Refresh the printer list.");
        var media = new List<PdfPrinterMedia>(driver.PaperCodes.Count);
        PdfPrinterMedia? defaultMedia = null;
        for (int index = 0; index < driver.PaperCodes.Count; index++) {
            string id = driver.PaperCodes[index].ToString(CultureInfo.InvariantCulture);
            if (media.Any(item => item.Id == id)) continue;
            (int width, int height) = driver.PaperSizes[index];
            string name = driver.PaperNames[index].Trim();
            var item = new PdfPrinterMedia(id, name.Length == 0 ? id : name,
                width > 0 && height > 0 ? PdfPrinterMediaNames.Portrait(width * 72D / 254D, height * 72D / 254D) : null);
            media.Add(item);
            if (driver.PaperCodes[index] == unchecked((ushort)driver.DefaultPaperCode)) defaultMedia = item;
        }
        return new PdfPrinterCapabilities {
            PrinterName = printerName,
            Media = media,
            DefaultMedia = defaultMedia,
            PaperSources = driver.PaperSources,
            SupportsDuplex = driver.Duplex < 0 ? null : driver.Duplex == 1,
            SupportsColor = driver.ColorDevice < 0 ? null : driver.ColorDevice == 1,
            SupportsCollation = driver.Collate < 0 ? null : driver.Collate == 1,
            MaximumCopies = driver.Copies <= 0 ? null : driver.Copies,
            Resolutions = driver.Resolutions.Where(resolution => resolution.X > 0 && resolution.Y > 0)
                .Select(resolution => new PdfPrinterResolution(resolution.X, resolution.Y)).Distinct().ToArray()
        };
    }

    /// <summary>Chooses driver copies only when the driver collates them and no paper change resets the device mid-job.</summary>
    internal static short DriverCopies(int requestedCopies, PdfPrinterCapabilities capabilities, bool singlePaper) =>
        requestedCopies > 1 && singlePaper && capabilities.SupportsCollation == true && capabilities.MaximumCopies >= requestedCopies
            ? checked((short)requestedCopies) : (short)1;
}
