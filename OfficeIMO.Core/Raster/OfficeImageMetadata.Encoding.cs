using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeImageMetadata {
    /// <summary>Creates an independent metadata copy containing only profile families carried by the destination format.</summary>
    /// <param name="format">The destination image container.</param>
    /// <param name="omittedProfiles">Profile families present in this metadata that the destination cannot carry.</param>
    /// <returns>An independent metadata copy suitable for the destination's supported profile families.</returns>
    /// <remarks>Use this projection when re-encoding pixels to another format. Lossless <see cref="Apply"/> remains strict about unsupported profiles. TIFF-relative opaque notes cannot be relocated to a new Exif-bearing container, including a newly encoded TIFF, until explicitly removed or replaced. An output with no Exif carrier omits the entire Exif family and reports that omission.</remarks>
    public OfficeImageMetadata PrepareForEncoding(OfficeImageFormat format, out OfficeImageMetadataProfileKinds omittedProfiles) {
        OfficeImageMetadataProfileKinds supported = format switch {
            OfficeImageFormat.Jpeg or OfficeImageFormat.Tiff => OfficeImageMetadataProfileKinds.Exif | OfficeImageMetadataProfileKinds.Xmp | OfficeImageMetadataProfileKinds.Icc | OfficeImageMetadataProfileKinds.Iptc,
            OfficeImageFormat.Png or OfficeImageFormat.Webp => OfficeImageMetadataProfileKinds.Exif | OfficeImageMetadataProfileKinds.Xmp | OfficeImageMetadataProfileKinds.Icc,
            OfficeImageFormat.Gif => OfficeImageMetadataProfileKinds.Xmp | OfficeImageMetadataProfileKinds.Icc,
            OfficeImageFormat.Bmp => OfficeImageMetadataProfileKinds.Icc,
            OfficeImageFormat.PortableMap or OfficeImageFormat.Tga or OfficeImageFormat.Icon => OfficeImageMetadataProfileKinds.None,
            _ => throw new ArgumentOutOfRangeException(nameof(format), "The destination is not a supported managed raster container.")
        };
        OfficeImageMetadataProfileKinds present = (HasExifProfile ? OfficeImageMetadataProfileKinds.Exif : 0) |
            (_xmp != null ? OfficeImageMetadataProfileKinds.Xmp : 0) |
            (_icc != null ? OfficeImageMetadataProfileKinds.Icc : 0) |
            (_iptc != null ? OfficeImageMetadataProfileKinds.Iptc : 0);
        omittedProfiles = present & ~supported;
        if (RequiresOriginalTiffContainer && (supported & OfficeImageMetadataProfileKinds.Exif) != 0) {
            throw new NotSupportedException("TIFF-relative maker-note offsets cannot be moved to a newly encoded container. Explicitly remove or replace the maker-note field before re-encoding, or use Apply with the original TIFF for lossless edits.");
        }
        OfficeImageMetadata result = Clone();
        if (format == OfficeImageFormat.Gif) result.ResolutionUnits = OfficeImageResolutionUnit.AspectRatio;
        if ((supported & OfficeImageMetadataProfileKinds.Exif) == 0) result.ClearExif();
        if ((supported & OfficeImageMetadataProfileKinds.Xmp) == 0) result._xmp = null;
        if ((supported & OfficeImageMetadataProfileKinds.Icc) == 0) result._icc = null;
        if ((supported & OfficeImageMetadataProfileKinds.Iptc) == 0) result._iptc = null;
        return result;
    }
}
