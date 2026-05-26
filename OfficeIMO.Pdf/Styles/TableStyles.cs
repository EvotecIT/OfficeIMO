namespace OfficeIMO.Pdf;

/// <summary>
/// Friendly presets for common table appearances.
/// </summary>
public static class TableStyles {
    /// <summary>
    /// Word table style names currently mapped to PDF presets.
    /// </summary>
    public static IReadOnlyList<string> SupportedWordStyleNames { get; } = Array.AsReadOnly(new[] {
        "TableGrid",
        "PlainTable1",
        "GridTable1Light",
        "ListTable1Light"
    });

    /// <summary>
    /// Light preset: report-friendly header fill, soft grid, comfortable padding, and gentle row striping.
    /// </summary>
    public static PdfTableStyle Light() => new PdfTableStyle {
        HeaderFill = PdfColor.FromRgb(32, 76, 120),
        HeaderTextColor = PdfColor.White,
        TextColor = PdfColor.FromRgb(31, 41, 55),
        RowStripeFill = PdfColor.FromRgb(248, 250, 252),
        BorderColor = PdfColor.FromRgb(210, 218, 226),
        BorderWidth = 0.5,
        CellPaddingX = 6,
        CellPaddingY = 5
    };

    /// <summary>
    /// Minimal preset: grid only with no header or row fills.
    /// </summary>
    public static PdfTableStyle Minimal() => new PdfTableStyle {
        BorderColor = PdfColor.FromRgb(210, 210, 210),
        BorderWidth = 0.5
    };

    /// <summary>
    /// Light preset with automatic right alignment for numeric-looking values.
    /// </summary>
    public static PdfTableStyle RightAlignedNumbers() {
        var t = Light();
        t.RightAlignNumeric = true;
        return t;
    }

    /// <summary>
    /// Creates a PDF table style from a Word table style name supported by OfficeIMO.Pdf.
    /// </summary>
    public static PdfTableStyle FromWordTableStyle(string styleName) {
        if (TryFromWordTableStyle(styleName, out PdfTableStyle? style)) {
            return style!;
        }

        throw new ArgumentException(
            $"Unsupported Word table style '{styleName}'. Supported styles: {string.Join(", ", SupportedWordStyleNames)}.",
            nameof(styleName));
    }

    /// <summary>
    /// Tries to create a PDF table style from a Word table style name supported by OfficeIMO.Pdf.
    /// </summary>
    public static bool TryFromWordTableStyle(string styleName, out PdfTableStyle? style) {
        Guard.NotNull(styleName, nameof(styleName));

        string normalized = NormalizeWordTableStyleName(styleName);
        if (string.Equals(normalized, "TableGrid", StringComparison.OrdinalIgnoreCase)) {
            style = TableGrid();
            return true;
        }

        if (string.Equals(normalized, "PlainTable1", StringComparison.OrdinalIgnoreCase)) {
            style = PlainTable1();
            return true;
        }

        if (string.Equals(normalized, "GridTable1Light", StringComparison.OrdinalIgnoreCase)) {
            style = GridTable1Light();
            return true;
        }

        if (string.Equals(normalized, "ListTable1Light", StringComparison.OrdinalIgnoreCase)) {
            style = ListTable1Light();
            return true;
        }

        style = null;
        return false;
    }

    /// <summary>
    /// Word-like Table Grid preset: a plain neutral grid with no shading.
    /// </summary>
    public static PdfTableStyle TableGrid() => new PdfTableStyle {
        HeaderFill = null,
        HeaderTextColor = PdfColor.Black,
        TextColor = PdfColor.FromRgb(25, 25, 25),
        FooterFill = null,
        RowStripeFill = null,
        BorderColor = PdfColor.FromRgb(191, 191, 191),
        BorderWidth = 0.5,
        CellPaddingX = 5,
        CellPaddingY = 4,
        HeaderRowCount = 1,
        FooterRowCount = 0
    };

    /// <summary>
    /// Word-like Plain Table 1 preset: text-only table flow with no visible grid or shading.
    /// </summary>
    public static PdfTableStyle PlainTable1() => new PdfTableStyle {
        HeaderFill = null,
        HeaderTextColor = PdfColor.Black,
        TextColor = PdfColor.FromRgb(25, 25, 25),
        FooterFill = null,
        RowStripeFill = null,
        BorderColor = null,
        BorderWidth = 0,
        CellPaddingX = 5,
        CellPaddingY = 4,
        HeaderRowCount = 1,
        FooterRowCount = 0
    };

    /// <summary>
    /// Word-like Grid Table 1 Light preset: a light neutral grid with a slightly stronger header separator.
    /// </summary>
    public static PdfTableStyle GridTable1Light() => new PdfTableStyle {
        HeaderFill = null,
        HeaderTextColor = PdfColor.Black,
        TextColor = PdfColor.FromRgb(25, 25, 25),
        FooterFill = null,
        RowStripeFill = null,
        BorderColor = PdfColor.FromRgb(217, 217, 217),
        BorderWidth = 0.45,
        HeaderSeparatorColor = PdfColor.FromRgb(127, 127, 127),
        HeaderSeparatorWidth = 0.8,
        FooterSeparatorColor = PdfColor.FromRgb(127, 127, 127),
        FooterSeparatorWidth = 0.8,
        CellPaddingX = 5,
        CellPaddingY = 5,
        HeaderRowCount = 1,
        FooterRowCount = 0
    };

    /// <summary>
    /// Word-like List Table 1 Light preset: no full grid, a strong header separator, and soft row separators.
    /// </summary>
    public static PdfTableStyle ListTable1Light() => new PdfTableStyle {
        HeaderFill = null,
        HeaderTextColor = PdfColor.Black,
        TextColor = PdfColor.FromRgb(25, 25, 25),
        FooterFill = null,
        RowStripeFill = null,
        BorderColor = null,
        BorderWidth = 0,
        RowSeparatorColor = PdfColor.FromRgb(224, 224, 224),
        RowSeparatorWidth = 0.45,
        HeaderSeparatorColor = PdfColor.Black,
        HeaderSeparatorWidth = 0.8,
        FooterSeparatorColor = PdfColor.Black,
        FooterSeparatorWidth = 0.8,
        CellPaddingX = 4,
        CellPaddingY = 6,
        HeaderRowCount = 1,
        FooterRowCount = 0
    };

    private static string NormalizeWordTableStyleName(string styleName) {
        string trimmed = styleName.Trim();
        if (trimmed.Length == 0) {
            return string.Empty;
        }

        var sb = new StringBuilder(trimmed.Length);
        for (int i = 0; i < trimmed.Length; i++) {
            char ch = trimmed[i];
            if (char.IsWhiteSpace(ch) || ch == '-' || ch == '_') {
                continue;
            }

            sb.Append(ch);
        }

        return sb.ToString();
    }
}
