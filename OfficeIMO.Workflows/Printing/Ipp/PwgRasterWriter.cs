using System.Text;

namespace OfficeIMO.Workflows;

/// <summary>Pixel layouts the PWG raster writer accepts (pwg-raster-document-type-supported keywords).</summary>
internal enum PwgRasterColorSpace {
    /// <summary>sgray_8: one byte per pixel, 0 is black and 255 is white.</summary>
    Gray8,
    /// <summary>srgb_8: three bytes per pixel in red, green, blue order.</summary>
    Rgb8
}

/// <summary>One sheet side of chunky 8-bit pixels at a device resolution.</summary>
internal sealed class PwgRasterPage {
    public PwgRasterPage(int width, int height, int resolutionX, int resolutionY, PwgRasterColorSpace colorSpace, byte[] pixels) {
        if (width is < 1 or > 65536) throw new ArgumentOutOfRangeException(nameof(width));
        if (height is < 1 or > 65536) throw new ArgumentOutOfRangeException(nameof(height));
        if (resolutionX is < 1 or > 9600) throw new ArgumentOutOfRangeException(nameof(resolutionX));
        if (resolutionY is < 1 or > 9600) throw new ArgumentOutOfRangeException(nameof(resolutionY));
        if (!OfficeRuntimeCompatibility.IsDefined(colorSpace)) throw new ArgumentOutOfRangeException(nameof(colorSpace));
        ArgumentNullException.ThrowIfNull(pixels);
        if ((long)width * height * (colorSpace == PwgRasterColorSpace.Rgb8 ? 3 : 1) != pixels.LongLength)
            throw new ArgumentException("The pixel buffer does not match the page dimensions.", nameof(pixels));
        Width = width;
        Height = height;
        ResolutionX = resolutionX;
        ResolutionY = resolutionY;
        ColorSpace = colorSpace;
        Pixels = pixels;
    }

    public int Width { get; }
    public int Height { get; }
    public int ResolutionX { get; }
    public int ResolutionY { get; }
    public PwgRasterColorSpace ColorSpace { get; }
    /// <summary>Rows from top to bottom, without padding.</summary>
    public byte[] Pixels { get; }
    /// <summary>PWG media size name, such as iso_a4_210x297mm.</summary>
    public string? PageSizeName { get; init; }
    /// <summary>media-type keyword, such as stationery.</summary>
    public string? MediaType { get; init; }
    /// <summary>Marks the page as part of a two-sided job.</summary>
    public bool Duplex { get; init; }
    /// <summary>Short-edge binding for two-sided jobs.</summary>
    public bool Tumble { get; init; }
    /// <summary>-1 when the pixels were mirrored across the feed direction (each row reversed), otherwise 1.</summary>
    public int CrossFeedTransform { get; init; } = 1;
    /// <summary>-1 when the pixels were mirrored along the feed direction (row order reversed), otherwise 1.</summary>
    public int FeedTransform { get; init; } = 1;

    internal int BytesPerPixel => ColorSpace == PwgRasterColorSpace.Rgb8 ? 3 : 1;
}

/// <summary>
/// Writes PWG Raster (PWG 5102.4): the RaS2 sync word, then a 1796-byte header and compressed rows for each page.
/// </summary>
internal sealed class PwgRasterWriter {
    internal const int HeaderSize = 1796;
    private const uint ColorSpaceSGray = 18;
    private const uint ColorSpaceSrgb = 19;

    private readonly Stream _output;
    private readonly int _totalPages;
    private bool _syncWritten;

    /// <param name="output">Destination stream; it is not disposed.</param>
    /// <param name="totalPageCount">Number of pages that will be written, or 0 when unknown.</param>
    public PwgRasterWriter(Stream output, int totalPageCount = 0) {
        _output = output ?? throw new ArgumentNullException(nameof(output));
        if (totalPageCount < 0) throw new ArgumentOutOfRangeException(nameof(totalPageCount));
        _totalPages = totalPageCount;
    }

    public void WritePage(PwgRasterPage page) => WriteEncodedPage(EncodePage(page));

    /// <summary>Encodes one page's header and compressed rows, so a repeated page can be written again without encoding it twice.</summary>
    public byte[] EncodePage(PwgRasterPage page) {
        ArgumentNullException.ThrowIfNull(page);
        using var encoded = new MemoryStream();
        encoded.Write(Header(page), 0, HeaderSize);
        WriteRows(page, encoded);
        return encoded.ToArray();
    }

    /// <summary>Writes a page produced by <see cref="EncodePage"/>, after the sync word when it is the first page.</summary>
    public void WriteEncodedPage(byte[] encoded) {
        ArgumentNullException.ThrowIfNull(encoded);
        if (!_syncWritten) {
            _output.Write(new[] { (byte)'R', (byte)'a', (byte)'S', (byte)'2' }, 0, 4);
            _syncWritten = true;
        }
        _output.Write(encoded, 0, encoded.Length);
    }

    private byte[] Header(PwgRasterPage page) {
        var header = new byte[HeaderSize];
        Text(header, 0, "PwgRaster"); // PWG 5102.4 section 4.3.1.1: the fixed PwgRaster field value
        if (page.MediaType is { } mediaType) Text(header, 128, mediaType);
        UInt(header, 272, page.Duplex ? 1u : 0u);
        UInt(header, 276, (uint)page.ResolutionX);
        UInt(header, 280, (uint)page.ResolutionY);
        UInt(header, 340, 1); // NumCopies: copies are a job attribute, not repeated pages
        UInt(header, 352, (uint)Math.Round(page.Width * 72.0 / page.ResolutionX)); // PageSize in points
        UInt(header, 356, (uint)Math.Round(page.Height * 72.0 / page.ResolutionY));
        UInt(header, 368, page.Tumble ? 1u : 0u);
        UInt(header, 372, (uint)page.Width);
        UInt(header, 376, (uint)page.Height);
        UInt(header, 384, 8); // BitsPerColor
        UInt(header, 388, (uint)(8 * page.BytesPerPixel));
        UInt(header, 392, (uint)(page.Width * page.BytesPerPixel));
        UInt(header, 396, 0); // ColorOrder: chunky
        UInt(header, 400, page.ColorSpace == PwgRasterColorSpace.Rgb8 ? ColorSpaceSrgb : ColorSpaceSGray);
        UInt(header, 420, (uint)page.BytesPerPixel); // NumColors
        UInt(header, 452, (uint)_totalPages);
        UInt(header, 456, unchecked((uint)page.CrossFeedTransform));
        UInt(header, 460, unchecked((uint)page.FeedTransform));
        UInt(header, 472, (uint)page.Width); // ImageBoxRight
        UInt(header, 476, (uint)page.Height); // ImageBoxBottom
        UInt(header, 480, 0x00FFFFFF); // AlternatePrimary: white
        if (page.PageSizeName is { } sizeName) Text(header, 1732, sizeName);
        return header;
    }

    /// <summary>Line-repeat counts with run-length pixel packets (PWG 5102.4 section 4.3.1.1).</summary>
    private static void WriteRows(PwgRasterPage page, Stream output) {
        int stride = page.Width * page.BytesPerPixel;
        byte[] pixels = page.Pixels;
        using var line = new MemoryStream(stride + stride / 64 + 16);
        int row = 0;
        while (row < page.Height) {
            int repeat = 1;
            while (repeat < 256 && row + repeat < page.Height && RowsEqual(pixels, row * stride, (row + repeat) * stride, stride)) repeat++;
            line.SetLength(0);
            line.WriteByte((byte)(repeat - 1));
            EncodeRow(line, pixels, row * stride, page.Width, page.BytesPerPixel);
            output.Write(line.GetBuffer(), 0, (int)line.Length);
            row += repeat;
        }
    }

    private static void EncodeRow(MemoryStream line, byte[] pixels, int offset, int width, int bpp) {
        int x = 0;
        while (x < width) {
            int run = 1;
            while (run < 128 && x + run < width && PixelsEqual(pixels, offset + x * bpp, offset + (x + run) * bpp, bpp)) run++;
            if (run > 1) {
                line.WriteByte((byte)(run - 1));
                line.Write(pixels, offset + x * bpp, bpp);
                x += run;
                continue;
            }
            // Gather distinct pixels until a repeated pair starts or the 128-pixel packet limit.
            int start = x;
            int count = 1;
            x++;
            while (x < width && count < 128 && !(x + 1 < width && PixelsEqual(pixels, offset + x * bpp, offset + (x + 1) * bpp, bpp))) {
                count++;
                x++;
            }
            line.WriteByte(count == 1 ? (byte)0 : (byte)(257 - count));
            line.Write(pixels, offset + start * bpp, count * bpp);
        }
    }

    private static bool RowsEqual(byte[] data, int a, int b, int length) {
        for (int i = 0; i < length; i++) {
            if (data[a + i] != data[b + i]) return false;
        }
        return true;
    }

    private static bool PixelsEqual(byte[] data, int a, int b, int bpp) {
        for (int i = 0; i < bpp; i++) {
            if (data[a + i] != data[b + i]) return false;
        }
        return true;
    }

    private static void UInt(byte[] header, int offset, uint value) {
        header[offset] = (byte)(value >> 24);
        header[offset + 1] = (byte)(value >> 16);
        header[offset + 2] = (byte)(value >> 8);
        header[offset + 3] = (byte)value;
    }

    private static void Text(byte[] header, int offset, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        Buffer.BlockCopy(bytes, 0, header, offset, Math.Min(bytes.Length, 63)); // 64-byte field, null-terminated
    }
}
