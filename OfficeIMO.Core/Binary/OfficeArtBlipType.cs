namespace OfficeIMO.Drawing.Binary;

/// <summary>Identifies the persistence format declared by an OfficeArt BLIP store entry.</summary>
public enum OfficeArtBlipType {
    /// <summary>The BLIP could not be read.</summary>
    Error = 0x00,

    /// <summary>The BLIP type is unknown.</summary>
    Unknown = 0x01,

    /// <summary>Enhanced Metafile image data.</summary>
    Emf = 0x02,

    /// <summary>Windows Metafile image data.</summary>
    Wmf = 0x03,

    /// <summary>Macintosh PICT image data.</summary>
    Pict = 0x04,

    /// <summary>JPEG image data.</summary>
    Jpeg = 0x05,

    /// <summary>PNG image data.</summary>
    Png = 0x06,

    /// <summary>Device-independent bitmap image data.</summary>
    Dib = 0x07,

    /// <summary>TIFF image data.</summary>
    Tiff = 0x11,

    /// <summary>CMYK or YCCK JPEG image data.</summary>
    CmykJpeg = 0x12
}

/// <summary>Identifies where an OfficeArt BLIP record was found.</summary>
public enum OfficeArtBlipStorage {
    /// <summary>No BLIP record was available.</summary>
    None,

    /// <summary>The BLIP record was embedded in its FBSE record.</summary>
    Embedded,

    /// <summary>The BLIP record was resolved through the FBSE delay-stream offset.</summary>
    Delayed
}
