using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>How a print job's sheets are represented.</summary>
public enum PdfPrintContentKind {
    /// <summary>Source pages are placed as vector PDF content on each sheet.</summary>
    Vector,
    /// <summary>Each sheet is a raster image, because the source cannot be reproduced as vector content.</summary>
    Raster
}

/// <summary>Resource limits for composing a print job.</summary>
public sealed class PdfPrintJobOptions {
    /// <summary>Maximum physical sheets in one job. Defaults to 10,000.</summary>
    public int MaximumSheets { get; set; } = 10_000;
    /// <summary>Maximum bytes of composed job content. Defaults to 256 MiB.</summary>
    public long MaximumContentBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Sheet resolution used only when the job content must be raster. Defaults to 300 DPI.</summary>
    public double RasterContentDpi { get; set; } = 300;
    /// <summary>Maximum pixels in one source-page or sheet raster, up to the shared decoder limit of 50 million.</summary>
    public long MaximumPixelsPerImage { get; set; } = 36_000_000;
}

/// <summary>Print content composed from a sheet plan, detached from later source-file or workspace changes.</summary>
/// <remarks>Preview and every printer transport use the same content. Transports that cannot send PDF rasterize it one sheet at a time.</remarks>
public sealed class PdfPrintJob {
    private readonly byte[] _content;
    private readonly object _renderLock = new();
    private PdfDocument? _contentDocument;

    internal PdfPrintJob(string sourcePath, PdfPrintPlan plan, byte[] content, PdfPrintContentKind contentKind,
        IReadOnlyList<string> diagnostics, long maximumPixelsPerImage, long maximumContentBytes) {
        SourcePath = sourcePath; Plan = plan; _content = content; ContentKind = contentKind;
        Diagnostics = Array.AsReadOnly(diagnostics.Distinct().ToArray());
        MaximumPixelsPerImage = maximumPixelsPerImage;
        MaximumContentBytes = maximumContentBytes;
    }

    /// <summary>Original source identity; never used to reopen content during delivery.</summary>
    public string SourcePath { get; }
    /// <summary>Plan the content was composed from; one content page per sheet.</summary>
    public PdfPrintPlan Plan { get; }
    /// <summary>Whether sheets carry vector PDF content or raster images.</summary>
    public PdfPrintContentKind ContentKind { get; }
    /// <summary>Notes from composition, such as why vector content was unavailable.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
    /// <summary>Length of the composed PDF content in bytes.</summary>
    public long ContentLength => _content.LongLength;
    internal long MaximumPixelsPerImage { get; }
    /// <summary>The composition byte limit, which also bounds content rasterized for delivery.</summary>
    internal long MaximumContentBytes { get; }

    /// <summary>Opens a read-only stream over the composed PDF, one page per sheet at its planned paper size.</summary>
    /// <remarks>Grayscale plans are applied when sheets are rasterized; the PDF itself keeps source colors.</remarks>
    public Stream OpenContent() => new MemoryStream(_content, writable: false);

    /// <summary>Composes a print job from an authenticated snapshot without reopening the source location.</summary>
    public static PdfPrintJob Create(PdfDocument document, PdfPrintPlanRequest request,
        PdfPrintJobOptions? options = null, CancellationToken cancellationToken = default) =>
        PdfPrintJobComposer.Compose(document, request, options ?? new PdfPrintJobOptions(), cancellationToken);

    /// <summary>Renders one sheet to PNG exactly as raster delivery produces it, including the plan's color mode.</summary>
    /// <param name="sheetNumber">One-based sheet number.</param>
    /// <param name="dpi">Resolution from 36 through 600 DPI.</param>
    /// <param name="cancellationToken">Cancels rendering.</param>
    public PdfPrintSheetImage RenderSheet(int sheetNumber, double dpi, CancellationToken cancellationToken = default) {
        (OfficeRasterImage raster, IReadOnlyList<string> diagnostics) = RenderRaster(sheetNumber, dpi, MaximumPixelsPerImage, cancellationToken);
        return new PdfPrintSheetImage(sheetNumber, OfficePngWriter.Encode(raster, cancellationToken), raster.Width, raster.Height, diagnostics);
    }

    /// <summary>Renders every sheet with the plan's color mode and wraps the images in a PDF, one sheet at a time.</summary>
    internal byte[] RasterizeToPdf(double maximumDpi, CancellationToken cancellationToken) =>
        PdfPrintJobComposer.WrapRasterSheets(Plan.Sheets, sheet => {
            double dpi = PdfPrintRaster.ResolveDpi(maximumDpi, maximumDpi, sheet.PaperSize.Width.ToPoints(), sheet.PaperSize.Height.ToPoints(), MaximumPixelsPerImage);
            return OfficePngWriter.Encode(RenderRaster(sheet.SheetNumber, dpi, MaximumPixelsPerImage, cancellationToken).Raster, cancellationToken);
        }, MaximumContentBytes, cancellationToken);

    internal (OfficeRasterImage Raster, IReadOnlyList<string> Diagnostics) RenderRaster(int sheetNumber, double dpi, long maximumPixels,
        CancellationToken cancellationToken) {
        if (sheetNumber < 1 || sheetNumber > Plan.Sheets.Count) throw new ArgumentOutOfRangeException(nameof(sheetNumber));
        if (!OfficeRuntimeCompatibility.IsFinite(dpi) || dpi < 36 || dpi > 600) throw new ArgumentOutOfRangeException(nameof(dpi));
        cancellationToken.ThrowIfCancellationRequested();
        PdfPageRenderResult page;
        lock (_renderLock) {
            _contentDocument ??= PdfDocument.LoadOwned(_content);
            page = _contentDocument.Render.DisplayPage(sheetNumber, new PdfPageDisplayOptions {
                Scale = dpi / 72, MaximumPixels = maximumPixels, MaximumOutputBytes = 256L * 1024 * 1024
            }, cancellationToken);
        }
        byte[] png = page.Bytes ?? throw new InvalidOperationException("Sheet " + sheetNumber + " could not be rendered: " + string.Join("; ", page.Diagnostics));
        if (!OfficeRasterImageDecoder.TryDecode(png, new OfficeRasterDecodeOptions {
            MaximumDecodedPixels = Math.Max(1L, (long)page.Width * page.Height), MaximumEncodedBytes = png.Length, CancellationToken = cancellationToken
        }, out OfficeRasterImage? raster, out _) || raster is null)
            throw new InvalidOperationException("A rendered sheet could not be decoded.");
        if (Plan.ColorMode == PdfPrintColorMode.Grayscale) PdfPrintRaster.ApplyGrayscale(raster, cancellationToken);
        return (raster, page.Diagnostics);
    }
}

/// <summary>One rendered sheet.</summary>
public sealed class PdfPrintSheetImage {
    private readonly byte[] _png;
    internal PdfPrintSheetImage(int sheetNumber, byte[] png, int width, int height, IReadOnlyList<string> diagnostics) {
        SheetNumber = sheetNumber; _png = png; Width = width; Height = height; Diagnostics = diagnostics;
    }
    /// <summary>One-based sheet number.</summary>
    public int SheetNumber { get; }
    /// <summary>Pixel width.</summary>
    public int Width { get; }
    /// <summary>Pixel height.</summary>
    public int Height { get; }
    /// <summary>Managed-renderer limitations reported for this sheet.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
    /// <summary>Returns a copy of the PNG bytes.</summary>
    public byte[] GetPng() => (byte[])_png.Clone();
}
