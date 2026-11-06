namespace OfficeIMO.Pdf;

/// <summary>Reuses the last small Indexed soft-mask PNG within one logical read.</summary>
internal sealed class PdfIndexedSoftMaskPngMemo {
    private const int MaximumPngBytes = 1024 * 1024;
    private (PdfStream Image, PdfStream Mask, int DecodeLimit, byte[] Palette, byte[] Png)? _last;

    internal byte[]? Find(PdfStream image, PdfStream mask, int decodeLimit, byte[] palette) {
        if (_last is not { } last || !ReferenceEquals(last.Image, image) ||
            !ReferenceEquals(last.Mask, mask) || last.DecodeLimit != decodeLimit ||
            !palette.AsSpan().SequenceEqual(last.Palette)) return null;
        return last.Png;
    }

    internal void Remember(PdfStream image, PdfStream mask, int decodeLimit, byte[] palette, byte[] png) {
        // The normalizer owns these buffers. Each extracted-image wrapper copies the PNG.
        if (png.Length <= MaximumPngBytes) _last = (image, mask, decodeLimit, palette, png);
    }
}
