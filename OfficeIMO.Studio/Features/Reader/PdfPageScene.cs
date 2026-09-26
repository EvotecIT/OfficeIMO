using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Reader;

/// <summary>
/// Retained OfficeIMO page drawing, interaction geometry, and rendering diagnostics used by the desktop surface.
/// </summary>
public sealed record PdfPageScene(
    int PageNumber,
    OfficeDrawing Drawing,
    PdfPageInteractionMap? Interactions,
    IReadOnlyList<string> Diagnostics,
    bool RequiresRasterFallback) {
    internal int ElementCount { get; } = CountElements(Drawing);

    /// <summary>True for a scanned page: it paints an image but has no extractable text to search or select.</summary>
    internal bool IsImageOnly { get; } = Interactions is { TextRegions.Count: 0 } && ContainsImage(Drawing);

    private static bool ContainsImage(OfficeDrawing drawing) => drawing.Elements.Any(element => element switch {
        OfficeDrawingImage or OfficeDrawingImagePattern => true,
        OfficeDrawingGroup group => ContainsImage(group.Drawing),
        OfficeDrawingEffectGroup effectGroup => ContainsImage(effectGroup.Drawing),
        _ => false
    });

    internal long EstimatedBytes { get; } = EstimateDrawingBytes(Drawing) +
        (long)(Interactions?.Regions.Count ?? 0) * 192L +
        Diagnostics.Sum(static diagnostic => (long)diagnostic.Length * sizeof(char));

    private static int CountElements(OfficeDrawing drawing) {
        int count = drawing.Elements.Count;
        foreach (OfficeDrawingElement element in drawing.Elements) {
            count += element switch {
                OfficeDrawingGroup group => CountElements(group.Drawing),
                OfficeDrawingEffectGroup effectGroup => CountElements(effectGroup.Drawing) +
                    (effectGroup.SoftMask is null ? 0 : CountElements(effectGroup.SoftMask.Drawing)),
                _ => 0
            };
        }

        return count;
    }

    private static long EstimateDrawingBytes(OfficeDrawing drawing) {
        long bytes = (long)drawing.Elements.Count * 256L +
            drawing.Fonts.Faces.Sum(static face => face.Data.LongLength);
        foreach (OfficeDrawingElement element in drawing.Elements) {
            bytes += element switch {
                OfficeDrawingImage image => EstimateImageBytes(image),
                OfficeDrawingImagePattern pattern => pattern.Bytes.LongLength,
                OfficeDrawingText text => (long)text.Text.Length * sizeof(char),
                OfficeDrawingGroup group => EstimateDrawingBytes(group.Drawing),
                OfficeDrawingEffectGroup effectGroup => EstimateDrawingBytes(effectGroup.Drawing) +
                    (effectGroup.SoftMask is null ? 0L : EstimateDrawingBytes(effectGroup.SoftMask.Drawing)),
                _ => 0L
            };
        }
        return bytes;
    }

    private static long EstimateImageBytes(OfficeDrawingImage image) {
        byte[] encoded = image.Bytes;
        long estimate = encoded.LongLength;
        if (OfficeImageReader.TryIdentifyByContent(encoded, null, out OfficeImageInfo info) &&
            info.Width > 0 && info.Height > 0)
            estimate += (long)info.Width * info.Height * 4L;
        return estimate;
    }
}
