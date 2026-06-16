namespace OfficeIMO.Rtf;

/// <summary>
/// Supported RTF picture payload formats.
/// </summary>
public enum RtfImageFormat {
    /// <summary>Unknown or unsupported picture format.</summary>
    Unknown,

    /// <summary>Portable Network Graphics.</summary>
    Png,

    /// <summary>JPEG image.</summary>
    Jpeg,

    /// <summary>Windows bitmap.</summary>
    Dib,

    /// <summary>Windows Metafile.</summary>
    Wmf,

    /// <summary>Enhanced Metafile.</summary>
    Emf
}
