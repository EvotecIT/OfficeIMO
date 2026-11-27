namespace OfficeIMO.PowerPoint {
    /// <summary>The content types PowerPoint image parts are written with, as the SDK's image part types declare them.</summary>
    internal static class PowerPointImageContentTypes {
        internal static string For(OfficeImageFormat format) => format switch {
            OfficeImageFormat.Png => "image/png",
            OfficeImageFormat.Jpeg => "image/jpeg",
            OfficeImageFormat.Gif => "image/gif",
            OfficeImageFormat.Bmp => "image/bmp",
            OfficeImageFormat.Tiff => "image/tiff",
            OfficeImageFormat.Svg => "image/svg+xml",
            OfficeImageFormat.Emf => "image/x-emf",
            OfficeImageFormat.Wmf => "image/x-wmf",
            OfficeImageFormat.Icon => "image/x-icon",
            OfficeImageFormat.Pcx => "image/x-pcx",
            OfficeImageFormat.Jpeg2000 => "image/jp2",
            _ => throw new NotSupportedException($"Image format {format} is not supported by PowerPoint image parts.")
        };
    }
}
