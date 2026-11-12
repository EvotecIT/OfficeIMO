namespace OfficeIMO.Pdf;

/// <summary>Shares the PNG packed for the last indexed soft-masked image across the pages of one read.</summary>
internal sealed class PdfIndexedSoftMaskPngMemo {
    private const int MaximumPngBytes = 1024 * 1024;
    private (PdfStream Image, PdfStream Mask, int DecodeLimit, byte[] Palette, byte[] Png)? _last;

    // Extracted images never write to their payload and hand out copies, so repeated placements
    // can share one buffer instead of each holding its own copy of the same PNG.
    internal byte[]? Find(PdfStream image, PdfStream mask, int decodeLimit, byte[] palette) {
        if (_last is not { } last || !ReferenceEquals(last.Image, image) ||
            !ReferenceEquals(last.Mask, mask) || last.DecodeLimit != decodeLimit ||
            !palette.SequenceEqual(last.Palette)) return null;
        return last.Png;
    }

    internal void Remember(PdfStream image, PdfStream mask, int decodeLimit, byte[] palette, byte[] png) {
        if (png.Length <= MaximumPngBytes) _last = (image, mask, decodeLimit, palette, png);
    }
}
