namespace OfficeIMO.Studio.Features.Reader;

/// <summary>
/// Immutable encoded page result retained by the bounded Studio render cache. Scale is pixels per PDF point;
/// the host supplies its presentation pixel scale and owns bitmap decoding.
/// </summary>
internal sealed record PdfRenderedPage(
    int PageNumber,
    double Scale,
    byte[] Bytes,
    int PixelWidth,
    int PixelHeight,
    TimeSpan Elapsed,
    IReadOnlyList<string> Diagnostics) {
    internal long ByteLength => Bytes.LongLength;
}
