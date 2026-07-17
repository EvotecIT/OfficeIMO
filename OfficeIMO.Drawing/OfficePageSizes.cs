namespace OfficeIMO.Drawing;

/// <summary>
/// Common physical page sizes for image/page composition.
/// </summary>
public static class OfficePageSizes {
    /// <summary>US Letter, 8.5 x 11 inches.</summary>
    public static OfficePageSize Letter => new OfficePageSize(8.5D, 11D);

    /// <summary>US Legal, 8.5 x 14 inches.</summary>
    public static OfficePageSize Legal => new OfficePageSize(8.5D, 14D);

    /// <summary>US Tabloid, 11 x 17 inches.</summary>
    public static OfficePageSize Tabloid => new OfficePageSize(11D, 17D);

    /// <summary>US Ledger, 17 x 11 inches.</summary>
    public static OfficePageSize Ledger => new OfficePageSize(17D, 11D);

    /// <summary>US Statement, 5.5 x 8.5 inches.</summary>
    public static OfficePageSize Statement => new OfficePageSize(5.5D, 8.5D);

    /// <summary>US Executive, 7.25 x 10.5 inches.</summary>
    public static OfficePageSize Executive => new OfficePageSize(7.25D, 10.5D);

    /// <summary>ISO A3, 297 x 420 millimeters.</summary>
    public static OfficePageSize A3 => OfficePageSize.FromMillimeters(297D, 420D);

    /// <summary>ISO A4, 210 x 297 millimeters.</summary>
    public static OfficePageSize A4 => OfficePageSize.FromMillimeters(210D, 297D);

    /// <summary>ISO A5, 148 x 210 millimeters.</summary>
    public static OfficePageSize A5 => OfficePageSize.FromMillimeters(148D, 210D);

    /// <summary>ISO A6, 105 x 148 millimeters.</summary>
    public static OfficePageSize A6 => OfficePageSize.FromMillimeters(105D, 148D);

    /// <summary>JIS B4, 257 x 364 millimeters.</summary>
    public static OfficePageSize B4Jis => OfficePageSize.FromMillimeters(257D, 364D);

    /// <summary>JIS B5, 182 x 257 millimeters.</summary>
    public static OfficePageSize B5Jis => OfficePageSize.FromMillimeters(182D, 257D);

    /// <summary>JIS B6, 128 x 182 millimeters.</summary>
    public static OfficePageSize B6Jis => OfficePageSize.FromMillimeters(128D, 182D);

    /// <summary>Japanese postcard, 100 x 148 millimeters.</summary>
    public static OfficePageSize JapanesePostcard => OfficePageSize.FromMillimeters(100D, 148D);

    /// <summary>Index card, 3 x 5 inches.</summary>
    public static OfficePageSize IndexCard => new OfficePageSize(3D, 5D);

    /// <summary>Billfold page, 3.75 x 6.75 inches.</summary>
    public static OfficePageSize Billfold => new OfficePageSize(3.75D, 6.75D);
}
