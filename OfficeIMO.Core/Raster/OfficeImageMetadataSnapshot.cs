namespace OfficeIMO.Drawing;

internal sealed class OfficeImageMetadataSnapshot {
    internal OfficeImageMetadataKinds Kinds { get; set; }
    internal bool HasColorRenderingMetadata { get; set; }
    internal bool HasDeviceMultichannel { get; set; }
    internal int JpegSamplePrecision { get; set; }
    internal int JpegFrameMarker { get; set; }
    internal bool HasDeviceCmyk { get; set; }
    internal bool HasOtherPngColorRenderingMetadata { get; set; }
    internal bool HasNonSrgbPngCalibration { get; set; }
    internal bool HasPngAnimation { get; set; }
    internal byte[]? Exif { get; set; }
    internal byte[]? Xmp { get; set; }
    internal byte[]? Icc { get; set; }
    internal bool HasDuplicateJpegExif { get; set; }
    internal bool HasExtendedJpegXmp { get; set; }
    internal bool HasDuplicateStandardJpegXmp { get; set; }
    internal bool ExifContainsResolution { get; set; }
    internal bool HasPhysicalResolution { get; set; }
    internal bool HasUnitlessResolution { get; set; }
    internal double? PhysicalDpiX { get; set; }
    internal double? PhysicalDpiY { get; set; }
}
