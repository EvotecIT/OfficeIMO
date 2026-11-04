namespace OfficeIMO.Drawing;

/// <summary>
/// Common physical page sizes for image/page composition.
/// </summary>
public static class OfficePageSizes {
    /// <summary>US Letter, 8.5 x 11 inches.</summary>
    public static OfficePageSize Letter => new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(8.5D), global::OfficeIMO.OfficeLength.Inches(11D));

    /// <summary>US Legal, 8.5 x 14 inches.</summary>
    public static OfficePageSize Legal => new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(8.5D), global::OfficeIMO.OfficeLength.Inches(14D));

    /// <summary>US Tabloid, 11 x 17 inches.</summary>
    public static OfficePageSize Tabloid => new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(11D), global::OfficeIMO.OfficeLength.Inches(17D));

    /// <summary>US Ledger, 17 x 11 inches.</summary>
    public static OfficePageSize Ledger => new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(17D), global::OfficeIMO.OfficeLength.Inches(11D));

    /// <summary>US Statement, 5.5 x 8.5 inches.</summary>
    public static OfficePageSize Statement => new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(5.5D), global::OfficeIMO.OfficeLength.Inches(8.5D));

    /// <summary>US Executive, 7.25 x 10.5 inches.</summary>
    public static OfficePageSize Executive => new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(7.25D), global::OfficeIMO.OfficeLength.Inches(10.5D));

    /// <summary>ISO A3, 297 x 420 millimeters.</summary>
    public static OfficePageSize A3 => new OfficePageSize(global::OfficeIMO.OfficeLength.Millimeters(297D), global::OfficeIMO.OfficeLength.Millimeters(420D));

    /// <summary>ISO A4, 210 x 297 millimeters.</summary>
    public static OfficePageSize A4 => new OfficePageSize(global::OfficeIMO.OfficeLength.Millimeters(210D), global::OfficeIMO.OfficeLength.Millimeters(297D));

    /// <summary>ISO A5, 148 x 210 millimeters.</summary>
    public static OfficePageSize A5 => new OfficePageSize(global::OfficeIMO.OfficeLength.Millimeters(148D), global::OfficeIMO.OfficeLength.Millimeters(210D));

    /// <summary>ISO A6, 105 x 148 millimeters.</summary>
    public static OfficePageSize A6 => new OfficePageSize(global::OfficeIMO.OfficeLength.Millimeters(105D), global::OfficeIMO.OfficeLength.Millimeters(148D));

    /// <summary>JIS B4, 257 x 364 millimeters.</summary>
    public static OfficePageSize B4Jis => new OfficePageSize(global::OfficeIMO.OfficeLength.Millimeters(257D), global::OfficeIMO.OfficeLength.Millimeters(364D));

    /// <summary>JIS B5, 182 x 257 millimeters.</summary>
    public static OfficePageSize B5Jis => new OfficePageSize(global::OfficeIMO.OfficeLength.Millimeters(182D), global::OfficeIMO.OfficeLength.Millimeters(257D));

    /// <summary>JIS B6, 128 x 182 millimeters.</summary>
    public static OfficePageSize B6Jis => new OfficePageSize(global::OfficeIMO.OfficeLength.Millimeters(128D), global::OfficeIMO.OfficeLength.Millimeters(182D));

    /// <summary>US C sheet, 17 x 22 inches.</summary>
    public static OfficePageSize CSheet => new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(17D), global::OfficeIMO.OfficeLength.Inches(22D));

    /// <summary>Number 9 envelope, 3.875 x 8.875 inches.</summary>
    public static OfficePageSize Envelope9 => new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(3.875D), global::OfficeIMO.OfficeLength.Inches(8.875D));

    /// <summary>Number 10 envelope, 4.125 x 9.5 inches.</summary>
    public static OfficePageSize Envelope10 => new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(4.125D), global::OfficeIMO.OfficeLength.Inches(9.5D));

    /// <summary>DL envelope, 110 x 220 millimeters.</summary>
    public static OfficePageSize EnvelopeDl => new OfficePageSize(global::OfficeIMO.OfficeLength.Millimeters(110D), global::OfficeIMO.OfficeLength.Millimeters(220D));

    /// <summary>C5 envelope, 162 x 229 millimeters.</summary>
    public static OfficePageSize EnvelopeC5 => new OfficePageSize(global::OfficeIMO.OfficeLength.Millimeters(162D), global::OfficeIMO.OfficeLength.Millimeters(229D));

    /// <summary>C4 envelope, 229 x 324 millimeters.</summary>
    public static OfficePageSize EnvelopeC4 => new OfficePageSize(global::OfficeIMO.OfficeLength.Millimeters(229D), global::OfficeIMO.OfficeLength.Millimeters(324D));

    /// <summary>B5 envelope, 176 x 250 millimeters.</summary>
    public static OfficePageSize EnvelopeB5 => new OfficePageSize(global::OfficeIMO.OfficeLength.Millimeters(176D), global::OfficeIMO.OfficeLength.Millimeters(250D));

    /// <summary>Monarch envelope, 3.875 x 7.5 inches.</summary>
    public static OfficePageSize EnvelopeMonarch => new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(3.875D), global::OfficeIMO.OfficeLength.Inches(7.5D));

    /// <summary>Japanese postcard, 100 x 148 millimeters.</summary>
    public static OfficePageSize JapanesePostcard => new OfficePageSize(global::OfficeIMO.OfficeLength.Millimeters(100D), global::OfficeIMO.OfficeLength.Millimeters(148D));

    /// <summary>Index card, 3 x 5 inches.</summary>
    public static OfficePageSize IndexCard => new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(3D), global::OfficeIMO.OfficeLength.Inches(5D));

    /// <summary>Billfold page, 3.75 x 6.75 inches.</summary>
    public static OfficePageSize Billfold => new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(3.75D), global::OfficeIMO.OfficeLength.Inches(6.75D));
}
