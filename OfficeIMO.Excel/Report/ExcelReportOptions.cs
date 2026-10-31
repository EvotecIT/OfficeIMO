namespace OfficeIMO.Excel;

/// <summary>Controls the identity, spacing and optional fill overrides of worksheet report blocks.</summary>
public sealed class ExcelReportOptions {
    /// <summary>Gets or sets the theme used by report content.</summary>
    public OfficeTheme Theme { get; set; } = OfficeTheme.Default;
    /// <summary>Gets or sets an optional section-header fill override.</summary>
    public OfficeColor? SectionHeaderFill { get; set; }
    /// <summary>Gets or sets an optional key-cell fill override.</summary>
    public OfficeColor? KeyFill { get; set; }
    /// <summary>Gets or sets the default number of rows between blocks.</summary>
    public int DefaultSpacingRows { get; set; } = 1;

    internal string SectionHeaderFillHex => (SectionHeaderFill ?? Theme.Palette.Accent1).Resolve(Theme).ToRgbHex();
    internal string KeyFillHex => (KeyFill ?? Theme.Palette.Background).Resolve(Theme).ToRgbHex();
    internal ExcelReportOptions Snapshot() {
        if (Theme == null) throw new ArgumentNullException(nameof(Theme));
        if (DefaultSpacingRows < 0) throw new ArgumentOutOfRangeException(nameof(DefaultSpacingRows));
        return (ExcelReportOptions)MemberwiseClone();
    }
}
