namespace OfficeIMO.Markup.PowerPoint;

internal sealed partial class OfficeMarkupPowerPointExporter {
    private static bool TryReadEmbeddedImage(string source, MarkupToPowerPointOptions options,
        out byte[] bytes, out OfficeImageInfo info) {
        if (options.AllowDataUriImages) {
            return OfficeImageReader.TryReadBase64DataUri(source, options.MaximumDataUriImageBytes, out bytes, out info);
        }
        bytes = Array.Empty<byte>();
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        return false;
    }

    private static bool IsStretchFit(string? fit) =>
        Normalize(fit ?? string.Empty) is "fill" or "stretch";
}
