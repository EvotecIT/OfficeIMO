namespace OfficeIMO.Pdf;

/// <summary>Shares the PNG packed for the last indexed soft-masked image across the pages of one read.</summary>
internal sealed class PdfIndexedSoftMaskPngMemo {
    private const int MaximumPngBytes = 1024 * 1024;
    private Entry? _last;

    // Extracted images never write to their payload and hand out copies, so repeated placements
    // can share one buffer instead of each holding its own copy of the same PNG.
    internal byte[]? Find(PdfStream image, PdfStream mask, int width, int height, int bitsPerComponent, int decodeLimit, byte[] palette) {
        if (_last is not { } last || !ReferenceEquals(last.Image, image) || !ReferenceEquals(last.Mask, mask) ||
            last.Width != width || last.Height != height || last.BitsPerComponent != bitsPerComponent ||
            last.DecodeLimit != decodeLimit || !palette.SequenceEqual(last.Palette)) return null;
        return last.Png;
    }

    internal void Remember(PdfStream image, PdfStream mask, int width, int height, int bitsPerComponent, int decodeLimit,
        byte[] palette, byte[] png) {
        if (png.Length <= MaximumPngBytes) _last = new Entry(image, mask, width, height, bitsPerComponent, decodeLimit, palette, png);
    }

    private sealed class Entry {
        internal Entry(PdfStream image, PdfStream mask, int width, int height, int bitsPerComponent, int decodeLimit, byte[] palette, byte[] png) {
            Image = image; Mask = mask; Width = width; Height = height; BitsPerComponent = bitsPerComponent;
            DecodeLimit = decodeLimit; Palette = palette; Png = png;
        }

        internal PdfStream Image { get; }
        internal PdfStream Mask { get; }
        internal int Width { get; }
        internal int Height { get; }
        internal int BitsPerComponent { get; }
        internal int DecodeLimit { get; }
        internal byte[] Palette { get; }
        internal byte[] Png { get; }
    }
}
