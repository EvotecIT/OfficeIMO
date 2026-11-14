using OfficeIMO.Drawing;

namespace OfficeIMO.Workflows;

/// <summary>
/// A forward-only PWG raster document that renders and encodes pages a few ahead on other cores while it is read, within the job's content byte limit.
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
    private readonly PrintSheetPrefetcher<EncodedPage> _ahead;
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
        // Preparing a page holds the rendered RGBA pixels and the device rows at the printer resolution.
        double renderDpi = (double)plan.RasterDpi / plan.RasterScale;
        long rendered = (long)(first.PaperSize.Width.ToPoints() / 72 * renderDpi) * (long)(first.PaperSize.Height.ToPoints() / 72 * renderDpi);
        long pageBytes = rendered * (4 + (long)plan.RasterScale * plan.RasterScale * (plan.ColorSpace == PwgRasterColorSpace.Rgb8 ? 3 : 1));
        _ahead = new PrintSheetPrefetcher<EncodedPage>(_pages.Length, PrintSheetPrefetcher<EncodedPage>.DepthFor(pageBytes, job.MaximumContentBytes),
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
        EncodedPage? prepared = _ahead.Take(index);
        EncodedPage page;
        if (_pages[index] is int sheet) {
            int key = PageKey(sheet, index);
            if (!_encoded.TryGet(key, out page)) {
                page = prepared ?? RenderPage(index, _token);
                if (_lastUse[key] > index) _encoded.Add(key, page, page.Bytes.LongLength);
            }
        } else {
            var blank = new byte[checked(_lastWidth * _lastHeight * BytesPerPixel)];
            for (int i = 0; i < blank.Length; i++) blank[i] = 0xFF;
            page = Encode(index, blank, _lastWidth, _lastHeight, null);
        }
        _lastWidth = page.Width;
        _lastHeight = page.Height;
        _writer.WriteEncodedPage(page.Bytes);
        _buffer.Position = 0;
    }

    private int BytesPerPixel => _plan.ColorSpace == PwgRasterColorSpace.Rgb8 ? 3 : 1;

    private (bool MirrorRows, bool ReverseRows) Transform(int index) =>
        _plan.Duplex && index % 2 == 1 ? BackSideTransform(_plan.SheetBack, _plan.Tumble) : (false, false);

    /// <summary>
    /// Renders and encodes page <paramref name="index"/>; safe on a worker thread, as it reads only the job, the plan and the
    /// writer's fixed page count. Sheets render at the plan's device resolution divided by its scale, and each pixel is repeated.
    /// </summary>
    private EncodedPage RenderPage(int index, CancellationToken token) {
        int sheet = _pages[index]!.Value;
        (bool mirrorRows, bool reverseRows) = Transform(index);
        OfficeRasterImage raster = _job.RenderRaster(sheet + 1, (double)_plan.RasterDpi / _plan.RasterScale, _job.MaximumPixelsPerImage, token).Raster;
        byte[] pixels = DevicePixels(raster, BytesPerPixel, mirrorRows, reverseRows, _plan.RasterScale, out int width, out int height, token);
        return Encode(index, pixels, width, height, _plan.MediaNames.Count > sheet ? _plan.MediaNames[sheet] : null);
    }

    private EncodedPage Encode(int index, byte[] pixels, int width, int height, string? mediaName) {
        (bool mirrorRows, bool reverseRows) = Transform(index);
        byte[] encoded = _writer.EncodePage(new PwgRasterPage(width, height, _plan.RasterDpi, _plan.RasterDpi, _plan.ColorSpace, pixels) {
            PageSizeName = mediaName, Duplex = _plan.Duplex, Tumble = _plan.Tumble,
            CrossFeedTransform = mirrorRows ? -1 : 1, FeedTransform = reverseRows ? -1 : 1
        });
        return new EncodedPage(encoded, width, height);
    }

    /// <summary>How a back side is mirrored for pwg-raster-document-sheet-back (PWG 5100.14): (mirror each row, reverse the row order).</summary>
    internal static (bool MirrorRows, bool ReverseRows) BackSideTransform(string? sheetBack, bool tumble) => sheetBack switch {
        "flipped" => tumble ? (true, false) : (false, true),
        "rotated" => tumble ? (false, false) : (true, true),
        "manual-tumble" => tumble ? (true, true) : (false, false),
        _ => (false, false)
    };

    /// <summary>
    /// Converts rendered RGBA pixels to device rows: composited on white, landscape sheets turned to the portrait feed direction
    /// (content top along the left edge, as orientation-requested landscape prints it), then the back-side mirroring applied, and
    /// each pixel repeated <paramref name="scale"/> times along both edges.
    /// </summary>
    internal static byte[] DevicePixels(OfficeRasterImage raster, int bytesPerPixel, bool mirrorRows, bool reverseRows, int scale,
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
        if (scale <= 1) return output;
        int smallWidth = width, smallHeight = height;
        width = checked(smallWidth * scale);
        height = checked(smallHeight * scale);
        int smallStride = smallWidth * bytesPerPixel, stride = width * bytesPerPixel;
        var enlarged = new byte[checked(stride * height)];
        for (int y = 0; y < smallHeight; y++) {
            if ((y & 63) == 0) token.ThrowIfCancellationRequested();
            int row = y * scale * stride;
            for (int x = 0, to = row; x < smallWidth; x++) {
                int from = y * smallStride + x * bytesPerPixel;
                for (int repeat = 0; repeat < scale; repeat++, to += bytesPerPixel) Buffer.BlockCopy(output, from, enlarged, to, bytesPerPixel);
            }
            for (int repeat = 1; repeat < scale; repeat++) Buffer.BlockCopy(enlarged, row, enlarged, row + repeat * stride, stride);
        }
        return enlarged;
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
