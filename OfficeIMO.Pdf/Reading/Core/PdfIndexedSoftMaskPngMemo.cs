namespace OfficeIMO.Pdf;

/// <summary>Reuses the last small Indexed soft-mask PNG within one logical read.</summary>
internal sealed class PdfIndexedSoftMaskPngMemo {
    private const int MaximumPngBytes = 1024 * 1024;
    private (PdfStream Image, PdfStream Mask, int DecodeLimit, byte[] Palette, byte[] Png)? _last;

    internal byte[]? Find(PdfStream image, PdfStream mask, int decodeLimit, byte[] palette) {
        if (_last is not { } last || !ReferenceEquals(last.Image, image) ||
            !ReferenceEquals(last.Mask, mask) || last.DecodeLimit != decodeLimit ||
            !palette.SequenceEqual(last.Palette)) return null;
        return (byte[])last.Png.Clone();
    }

    internal void Remember(PdfStream image, PdfStream mask, int decodeLimit, byte[] palette, byte[] png) {
        // Extracted images own the buffers they are given, so the memo keeps and hands out its own copies.
        if (png.Length <= MaximumPngBytes) _last = (image, mask, decodeLimit, palette, (byte[])png.Clone());
    }
}
