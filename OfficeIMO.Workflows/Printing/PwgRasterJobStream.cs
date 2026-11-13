using OfficeIMO.Drawing;

namespace OfficeIMO.Workflows;

/// <summary>
/// A forward-only PWG raster document that renders pages a few ahead on other cores while it is read, within the job's content byte limit.
/// Sheets repeated for software copies are encoded once and written again from the kept page, within the same limit.
/// </summary>
/// <remarks>The first page is rendered on construction, so rendering failures surface before any data reaches the printer.</remarks>
internal sealed class PwgRasterJobStream : Stream {
    private readonly PdfPrintJob _job;
    private readonly IppDeliveryPlan _plan;
    private readonly CancellationToken _token;
    private readonly int?[] _pages;
    private readonly MemoryStream _buffer = new();
    private readonly PwgRasterWriter _writer;
    private readonly PrintSheetCache<EncodedPage> _encoded;
    private readonly int[] _lastUse;
    private readonly PrintSheetPrefetcher<DevicePage> _ahead;
    private int _next;
    private int _lastWidth, _lastHeight;

    public PwgRasterJobStream(PdfPrintJob job, IppDeliveryPlan plan, CancellationToken token)
        : this(job, plan, job.MaximumContentBytes, token) { }

    internal PwgRasterJobStream(PdfPrintJob job, IppDeliveryPlan plan, long keptPageBytes, CancellationToken token) {
        _job = job;
        _plan = plan;
        _token = token;
        _pages = PageSequence(job.Plan.Sheets.Count, plan.SoftwareCopies, plan.Duplex);
        _writer = new PwgRasterWriter(_buffer, _pages.Length);
        _encoded = new PrintSheetCache<EncodedPage>(keptPageBytes);
        _lastUse = new int[job.Plan.Sheets.Count * 2];
        for (int index = 0; index < _pages.Length; index++)
            if (_pages[index] is int sheet) _lastUse[PageKey(sheet, index)] = index;
        // Only the first use of each page renders ahead; later uses come from the kept encoding, or render again if it did not fit.
        var firstUse = new HashSet<int>();
        var renderAhead = new bool[_pages.Length];
        for (int index = 0; index < _pages.Length; index++)
            if (_pages[index] is int sheet) renderAhead[index] = firstUse.Add(PageKey(sheet, index));
        PdfPrintSheet first = job.Plan.Sheets[0];
        long pageBytes = (long)(first.PaperSize.Width.ToPoints() / 72 * plan.RasterDpi) * (long)(first.PaperSize.Height.ToPoints() / 72 * plan.RasterDpi) *
            (4 + (plan.ColorSpace == PwgRasterColorSpace.Rgb8 ? 3 : 1));
        _ahead = new PrintSheetPrefetcher<DevicePage>(_pages.Length, PrintSheetPrefetcher<DevicePage>.DepthFor(pageBytes, job.MaximumContentBytes),
            (index, stop) => renderAhead[index] ? RenderPage(index, stop) : null, token);
        try {
            WriteNextPage();
        } catch {
            _ahead.Dispose();
            throw;
        }
    }

    /// <summary>Encoded pages written again instead of being rendered for another copy.</summary>
    internal int ReusedPages => _encoded.Reused;

    // A sheet's encoding depends on whether it lands on a back side, which mirrors its pixels.
    private int PageKey(int sheet, int index) => sheet * 2 + (_plan.Duplex && index % 2 == 1 ? 1 : 0);

    /// <summary>Sheet indexes for each page, repeating the sequence per copy; null is a blank back side that starts each duplex copy on a new sheet.</summary>
    internal static int?[] PageSequence(int sheets, int copies, bool duplex) {
        var pages = new List<int?>(sheets * copies + copies);
        for (int copy = 0; copy < copies; copy++) {
            for (int sheet = 0; sheet < sheets; sheet++) pages.Add(sheet);
            if (duplex && sheets % 2 != 0 && copy + 1 < copies) pages.Add(null);
        }
        return pages.ToArray();
    }

    public override int Read(byte[] buffer, int offset, int count) {
        while (_buffer.Position == _buffer.Length) {
            if (_next >= _pages.Length) return 0;
            _buffer.SetLength(0);
            WriteNextPage();
        }
        return _buffer.Read(buffer, offset, count);
    }

    private void WriteNextPage() {
        _token.ThrowIfCancellationRequested();
        int index = _next++;
        _buffer.SetLength(0);
        DevicePage? prepared = _ahead.Take(index);
        if (_pages[index] is int repeated && _encoded.TryGet(PageKey(repeated, index), out EncodedPage kept)) {
            _lastWidth = kept.Width;
            _lastHeight = kept.Height;
            _writer.WriteEncodedPage(kept.Bytes);
            _buffer.Position = 0;
            return;
        }
        (bool mirrorRows, bool reverseRows) = Transform(index);
        int bytesPerPixel = BytesPerPixel;
        byte[] pixels;
        int width, height;
        string? mediaName = null;
        if (_pages[index] is int sheet) {
            DevicePage page = prepared ?? RenderPage(index, _token);
            pixels = page.Pixels;
            width = page.Width;
            height = page.Height;
            mediaName = _plan.MediaNames.Count > sheet ? _plan.MediaNames[sheet] : null;
        } else {
            width = _lastWidth;
            height = _lastHeight;
            pixels = new byte[checked(width * height * bytesPerPixel)];
            for (int i = 0; i < pixels.Length; i++) pixels[i] = 0xFF;
        }
        _lastWidth = width;
        _lastHeight = height;
        byte[] encoded = _writer.EncodePage(new PwgRasterPage(width, height, _plan.RasterDpi, _plan.RasterDpi, _plan.ColorSpace, pixels) {
            PageSizeName = mediaName, Duplex = _plan.Duplex, Tumble = _plan.Tumble,
            CrossFeedTransform = mirrorRows ? -1 : 1, FeedTransform = reverseRows ? -1 : 1
        });
        if (_pages[index] is int rendered && _lastUse[PageKey(rendered, index)] > index) _encoded.Add(PageKey(rendered, index), new EncodedPage(encoded, width, height), encoded.LongLength);
        _writer.WriteEncodedPage(encoded);
        _buffer.Position = 0;
    }

    private int BytesPerPixel => _plan.ColorSpace == PwgRasterColorSpace.Rgb8 ? 3 : 1;

    private (bool MirrorRows, bool ReverseRows) Transform(int index) =>
        _plan.Duplex && index % 2 == 1 ? BackSideTransform(_plan.SheetBack, _plan.Tumble) : (false, false);

    /// <summary>Renders page <paramref name="index"/> to device rows; safe on a worker thread, as it reads only the job and plan.</summary>
    private DevicePage RenderPage(int index, CancellationToken token) {
        int sheet = _pages[index]!.Value;
        (bool mirrorRows, bool reverseRows) = Transform(index);
        OfficeRasterImage raster = _job.RenderRaster(sheet + 1, _plan.RasterDpi, _job.MaximumPixelsPerImage, token).Raster;
        byte[] pixels = DevicePixels(raster, BytesPerPixel, mirrorRows, reverseRows, out int width, out int height, token);
        return new DevicePage(pixels, width, height);
    }

    private sealed record DevicePage(byte[] Pixels, int Width, int Height);

    /// <summary>How a back side is mirrored for pwg-raster-document-sheet-back (PWG 5100.14): (mirror each row, reverse the row order).</summary>
    internal static (bool MirrorRows, bool ReverseRows) BackSideTransform(string? sheetBack, bool tumble) => sheetBack switch {
        "flipped" => tumble ? (true, false) : (false, true),
        "rotated" => tumble ? (false, false) : (true, true),
        "manual-tumble" => tumble ? (true, true) : (false, false),
        _ => (false, false)
    };

    /// <summary>
    /// Converts rendered RGBA pixels to device rows: composited on white, landscape sheets turned to the portrait feed direction
    /// (content top along the left edge, as orientation-requested landscape prints it), then the back-side mirroring applied.
    /// </summary>
    internal static byte[] DevicePixels(OfficeRasterImage raster, int bytesPerPixel, bool mirrorRows, bool reverseRows,
        out int width, out int height, CancellationToken token) {
        int sourceWidth = raster.Width, sourceHeight = raster.Height;
        bool rotate = sourceWidth > sourceHeight;
        width = rotate ? sourceHeight : sourceWidth;
        height = rotate ? sourceWidth : sourceHeight;
        byte[] source = raster.PixelBuffer;
        var output = new byte[checked(width * height * bytesPerPixel)];
        for (int y = 0; y < height; y++) {
            if ((y & 63) == 0) token.ThrowIfCancellationRequested();
            int deviceY = reverseRows ? height - 1 - y : y;
            for (int x = 0; x < width; x++) {
                // A counter-clockwise turn puts the source's top row on the device's left edge and its right column on the top edge.
                int sourceX = rotate ? sourceWidth - 1 - y : x;
                int sourceY = rotate ? x : y;
                int from = (sourceY * sourceWidth + sourceX) * 4;
                int alpha = source[from + 3];
                int r = (source[from] * alpha + 255 * (255 - alpha) + 127) / 255;
                int g = (source[from + 1] * alpha + 255 * (255 - alpha) + 127) / 255;
                int b = (source[from + 2] * alpha + 255 * (255 - alpha) + 127) / 255;
                int deviceX = mirrorRows ? width - 1 - x : x;
                int to = (deviceY * width + deviceX) * bytesPerPixel;
                if (bytesPerPixel == 3) {
                    output[to] = (byte)r; output[to + 1] = (byte)g; output[to + 2] = (byte)b;
                } else {
                    output[to] = (byte)((r * 77 + g * 150 + b * 29 + 128) >> 8);
                }
            }
        }
        return output;
    }

    private sealed record EncodedPage(byte[] Bytes, int Width, int Height);

    public override bool CanRead => true;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();
    public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
    public override void Flush() { }
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    protected override void Dispose(bool disposing) {
        if (disposing) {
            _ahead.Dispose();
            _buffer.Dispose();
        }
        base.Dispose(disposing);
    }
}
