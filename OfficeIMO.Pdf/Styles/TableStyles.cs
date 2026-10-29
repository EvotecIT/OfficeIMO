namespace OfficeIMO.Pdf;

/// <summary>
/// Friendly presets for common table appearances.
/// </summary>
public static class PdfTableStyles {
    private static readonly string[] CanonicalWordStyleNameValues = {
        "TableNormal",
        "TableGrid",
        "TableGridLight",
        "PlainTable1",
        "GridTable1Light",
        "GridTable1LightAccent1",
        "GridTable1LightAccent2",
        "GridTable1LightAccent3",
        "GridTable1LightAccent4",
        "GridTable1LightAccent5",
        "GridTable1LightAccent6",
        "ListTable1Light",
        "ListTable1LightAccent1",
        "ListTable1LightAccent2",
        "ListTable1LightAccent3",
        "ListTable1LightAccent4",
        "ListTable1LightAccent5",
        "ListTable1LightAccent6"
    };

    private static readonly string[] WordStyleAliasNameValues = {
        "GridTableLight",
        "GridTable1Light-Accent1",
        "GridTable1Light-Accent2",
        "GridTable1Light-Accent3",
        "GridTable1Light-Accent4",
        "GridTable1Light-Accent5",
        "GridTable1Light-Accent6",
        "ListTable1Light-Accent1",
        "ListTable1Light-Accent2",
        "ListTable1Light-Accent3",
        "ListTable1Light-Accent4",
        "ListTable1Light-Accent5",
        "ListTable1Light-Accent6"
    };

    /// <summary>
    /// Canonical Word table style names callers can present without duplicate alias spellings.
    /// </summary>
    public static IReadOnlyList<string> CanonicalWordStyleNames { get; } = Array.AsReadOnly(CanonicalWordStyleNameValues);

    /// <summary>
    /// Word table style names and aliases currently mapped to PDF presets.
    /// </summary>
    public static IReadOnlyList<string> SupportedWordStyleNames { get; } = Array.AsReadOnly(CreateSupportedWordStyleNames());

    /// <summary>
    /// Light preset: report-friendly header fill, soft grid, comfortable padding, and gentle row striping.
    /// </summary>
    public static PdfTableStyle Light() => new PdfTableStyle {
        HeaderFill = global::OfficeIMO.OfficeColor.FromRgb(32, 76, 120),
        HeaderTextColor = global::OfficeIMO.OfficeColor.White,
        TextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
        RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
        BorderColor = global::OfficeIMO.OfficeColor.FromRgb(210, 218, 226),
        BorderWidth = 0.5,
        CellPaddingX = 6,
        CellPaddingY = 5,
        SpacingBefore = 4,
        SpacingAfter = 8
    };

    /// <summary>
    /// Minimal preset: grid only with no header or row fills.
    /// </summary>
    public static PdfTableStyle Minimal() => new PdfTableStyle {
        BorderColor = global::OfficeIMO.OfficeColor.FromRgb(210, 210, 210),
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
    /// Technical-document preset: compact readable rows, dark header fill, soft borders, and automatic column fitting for specifications and READMEs.
    /// </summary>
    public static PdfTableStyle TechnicalDocument() {
        PdfTableStyle style = CreateDocumentTable(Light(), 9.75, 9.75, 1.2, 6, 5, 6, 10);
        style.HeaderFill = global::OfficeIMO.OfficeColor.FromRgb(15, 23, 42);
        style.HeaderTextColor = global::OfficeIMO.OfficeColor.White;
        style.RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252);
        style.BorderColor = global::OfficeIMO.OfficeColor.FromRgb(203, 213, 225);
        style.RowSeparatorColor = global::OfficeIMO.OfficeColor.FromRgb(226, 232, 240);
        style.BorderWidth = 0.45;
        style.HeaderSeparatorWidth = 0.75;
        return style;
    }

    /// <summary>
    /// Compact preset: dense list-style table rhythm for notes, narrow reports, and space-constrained exports.
    /// </summary>
    public static PdfTableStyle Compact() {
        PdfTableStyle style = CreateDocumentTable(ListTable1Light(), 9, 9, 1.12, 4, 3, 4, 7);
        style.RowSeparatorColor = global::OfficeIMO.OfficeColor.FromRgb(226, 232, 240);
        return style;
    }

    /// <summary>
    /// Report preset: stronger header hierarchy, calm blue row striping, and automatic column fitting for polished summaries.
    /// </summary>
    public static PdfTableStyle Report() {
        PdfTableStyle style = CreateDocumentTable(Light(), 9.25, 9.25, 1.18, 5, 4, 6, 10);
        style.HeaderFill = global::OfficeIMO.OfficeColor.FromRgb(30, 64, 175);
        style.HeaderTextColor = global::OfficeIMO.OfficeColor.White;
        style.RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(239, 246, 255);
        style.BorderColor = global::OfficeIMO.OfficeColor.FromRgb(191, 219, 254);
        style.BorderWidth = 0.45;
        return style;
    }

    /// <summary>
    /// Creates a PDF table style from a Word table style name supported by OfficeIMO.Pdf.
    /// </summary>
    public static PdfTableStyle FromWordTableStyle(string styleName) {
        if (TryFromWordTableStyle(styleName, out PdfTableStyle? style)) {
            return style!;
        }

        throw CreateUnsupportedWordTableStyleException(styleName);
    }

    /// <summary>
    /// Tries to create a PDF table style from a Word table style name supported by OfficeIMO.Pdf.
    /// </summary>
    public static bool TryFromWordTableStyle(string styleName, out PdfTableStyle? style) {
        Guard.NotNull(styleName, nameof(styleName));

        string normalized = NormalizeWordTableStyleName(styleName);
        if (!TryGetCanonicalWordStyleNameFromNormalized(normalized, out string? canonicalStyleName)) {
            style = null;
            return false;
        }

        if (string.Equals(canonicalStyleName, "TableGrid", StringComparison.OrdinalIgnoreCase)) {
            style = TableGrid();
            return true;
        }

        if (string.Equals(canonicalStyleName, "TableGridLight", StringComparison.OrdinalIgnoreCase)) {
            style = TableGridLight();
            return true;
        }

        if (string.Equals(canonicalStyleName, "TableNormal", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(canonicalStyleName, "PlainTable1", StringComparison.OrdinalIgnoreCase)) {
            style = PlainTable1();
            return true;
        }

        if (string.Equals(canonicalStyleName, "GridTable1Light", StringComparison.OrdinalIgnoreCase)) {
            style = GridTable1Light();
            return true;
        }

        if (TryGetAccentNumber(canonicalStyleName!, "GridTable1LightAccent", out int gridAccent)) {
            style = GridTable1LightAccent(gridAccent);
            return true;
        }

        if (string.Equals(canonicalStyleName, "ListTable1Light", StringComparison.OrdinalIgnoreCase)) {
            style = ListTable1Light();
            return true;
        }

        if (TryGetAccentNumber(canonicalStyleName!, "ListTable1LightAccent", out int listAccent)) {
            style = ListTable1LightAccent(listAccent);
            return true;
        }

        style = null;
        return false;
    }

    /// <summary>
    /// Returns the canonical Word table style name for a supported style name or alias.
    /// </summary>
    public static string GetCanonicalWordStyleName(string styleName) {
        if (TryGetCanonicalWordStyleName(styleName, out string? canonicalStyleName)) {
            return canonicalStyleName!;
        }

        throw CreateUnsupportedWordTableStyleException(styleName);
    }

    /// <summary>
    /// Tries to return the canonical Word table style name for a supported style name or alias.
    /// </summary>
    public static bool TryGetCanonicalWordStyleName(string styleName, out string? canonicalStyleName) {
        Guard.NotNull(styleName, nameof(styleName));
        return TryGetCanonicalWordStyleNameFromNormalized(NormalizeWordTableStyleName(styleName), out canonicalStyleName);
    }

    /// <summary>
    /// Word-like Table Grid preset: a plain automatic-color grid with no shading.
    /// </summary>
    public static PdfTableStyle TableGrid() => new PdfTableStyle {
        HeaderFill = null,
        HeaderTextColor = global::OfficeIMO.OfficeColor.Black,
        TextColor = global::OfficeIMO.OfficeColor.FromRgb(25, 25, 25),
        FooterFill = null,
        RowStripeFill = null,
        BorderColor = global::OfficeIMO.OfficeColor.Black,
        BorderWidth = 0.5,
        CellPaddingX = 5,
        CellPaddingY = 4,
        HeaderBold = false,
        HeaderRowCount = 1,
        FooterRowCount = 0
    };

    /// <summary>
    /// Word-like Table Grid Light preset: a light neutral grid matching Word's built-in Grid Table Light style.
    /// </summary>
    public static PdfTableStyle TableGridLight() => new PdfTableStyle {
        HeaderFill = null,
        HeaderTextColor = global::OfficeIMO.OfficeColor.Black,
        TextColor = global::OfficeIMO.OfficeColor.FromRgb(25, 25, 25),
        FooterFill = null,
        RowStripeFill = null,
        BorderColor = global::OfficeIMO.OfficeColor.FromRgb(191, 191, 191),
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
        HeaderTextColor = global::OfficeIMO.OfficeColor.Black,
        TextColor = global::OfficeIMO.OfficeColor.FromRgb(25, 25, 25),
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
        HeaderTextColor = global::OfficeIMO.OfficeColor.Black,
        TextColor = global::OfficeIMO.OfficeColor.FromRgb(25, 25, 25),
        FooterFill = null,
        RowStripeFill = null,
        BorderColor = global::OfficeIMO.OfficeColor.FromRgb(217, 217, 217),
        BorderWidth = 0.45,
        HeaderSeparatorColor = global::OfficeIMO.OfficeColor.FromRgb(127, 127, 127),
        HeaderSeparatorWidth = 0.8,
        FooterSeparatorColor = global::OfficeIMO.OfficeColor.FromRgb(127, 127, 127),
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
        HeaderTextColor = global::OfficeIMO.OfficeColor.Black,
        TextColor = global::OfficeIMO.OfficeColor.FromRgb(25, 25, 25),
        FooterFill = null,
        RowStripeFill = null,
        BorderColor = null,
        BorderWidth = 0,
        RowSeparatorColor = global::OfficeIMO.OfficeColor.FromRgb(224, 224, 224),
        RowSeparatorWidth = 0.45,
        HeaderSeparatorColor = global::OfficeIMO.OfficeColor.Black,
        HeaderSeparatorWidth = 0.8,
        FooterSeparatorColor = global::OfficeIMO.OfficeColor.Black,
        FooterSeparatorWidth = 0.8,
        CellPaddingX = 4,
        CellPaddingY = 6,
        HeaderRowCount = 1,
        FooterRowCount = 0
    };

    private static PdfTableStyle GridTable1LightAccent(int accentNumber) {
        PdfTableStyle style = GridTable1Light();
        WordAccentColors colors = GetWordTableAccentColors(accentNumber);
        style.BorderColor = colors.Light;
        style.HeaderSeparatorColor = colors.Strong;
        style.FooterSeparatorColor = colors.Strong;
        return style;
    }

    private static PdfTableStyle ListTable1LightAccent(int accentNumber) {
        PdfTableStyle style = ListTable1Light();
        WordAccentColors colors = GetWordTableAccentColors(accentNumber);
        style.RowStripeFill = colors.Pale;
        style.HeaderSeparatorColor = colors.Strong;
        style.FooterSeparatorColor = colors.Strong;
        return style;
    }

    private readonly struct WordAccentColors {
        public WordAccentColors(global::OfficeIMO.OfficeColor light, global::OfficeIMO.OfficeColor strong, global::OfficeIMO.OfficeColor pale) {
            Light = light;
            Strong = strong;
            Pale = pale;
        }

        public global::OfficeIMO.OfficeColor Light { get; }
        public global::OfficeIMO.OfficeColor Strong { get; }
        public global::OfficeIMO.OfficeColor Pale { get; }
    }

    private static WordAccentColors GetWordTableAccentColors(int accentNumber) => accentNumber switch {
        1 => new WordAccentColors(global::OfficeIMO.OfficeColor.FromRgb(180, 198, 231), global::OfficeIMO.OfficeColor.FromRgb(142, 170, 219), global::OfficeIMO.OfficeColor.FromRgb(217, 226, 243)),
        2 => new WordAccentColors(global::OfficeIMO.OfficeColor.FromRgb(247, 202, 172), global::OfficeIMO.OfficeColor.FromRgb(244, 176, 131), global::OfficeIMO.OfficeColor.FromRgb(251, 228, 213)),
        3 => new WordAccentColors(global::OfficeIMO.OfficeColor.FromRgb(219, 219, 219), global::OfficeIMO.OfficeColor.FromRgb(201, 201, 201), global::OfficeIMO.OfficeColor.FromRgb(237, 237, 237)),
        4 => new WordAccentColors(global::OfficeIMO.OfficeColor.FromRgb(255, 229, 153), global::OfficeIMO.OfficeColor.FromRgb(255, 217, 102), global::OfficeIMO.OfficeColor.FromRgb(255, 242, 204)),
        5 => new WordAccentColors(global::OfficeIMO.OfficeColor.FromRgb(189, 214, 238), global::OfficeIMO.OfficeColor.FromRgb(156, 194, 229), global::OfficeIMO.OfficeColor.FromRgb(222, 234, 246)),
        6 => new WordAccentColors(global::OfficeIMO.OfficeColor.FromRgb(197, 224, 179), global::OfficeIMO.OfficeColor.FromRgb(168, 208, 141), global::OfficeIMO.OfficeColor.FromRgb(226, 239, 217)),
        _ => throw new ArgumentOutOfRangeException(nameof(accentNumber), "Word table accent number must be between 1 and 6.")
    };

    private static PdfTableStyle CreateDocumentTable(PdfTableStyle style, double fontSize, double headerFontSize, double lineHeight, double paddingX, double paddingY, double spacingBefore, double spacingAfter) {
        style.FontSize = fontSize;
        style.HeaderFontSize = headerFontSize;
        style.LineHeight = lineHeight;
        style.CellPaddingX = paddingX;
        style.CellPaddingY = paddingY;
        style.SpacingBefore = spacingBefore;
        style.SpacingAfter = spacingAfter;
        style.AutoFitColumns = true;
        return style;
    }

    private static string[] CreateSupportedWordStyleNames() {
        string[] names = new string[CanonicalWordStyleNameValues.Length + WordStyleAliasNameValues.Length];
        Array.Copy(CanonicalWordStyleNameValues, names, CanonicalWordStyleNameValues.Length);
        Array.Copy(WordStyleAliasNameValues, 0, names, CanonicalWordStyleNameValues.Length, WordStyleAliasNameValues.Length);
        return names;
    }

    private static bool TryGetCanonicalWordStyleNameFromNormalized(string normalized, out string? canonicalStyleName) {
        if (string.Equals(normalized, "TableNormal", StringComparison.OrdinalIgnoreCase)) {
            canonicalStyleName = "TableNormal";
            return true;
        }

        if (string.Equals(normalized, "TableGrid", StringComparison.OrdinalIgnoreCase)) {
            canonicalStyleName = "TableGrid";
            return true;
        }

        if (string.Equals(normalized, "TableGridLight", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(normalized, "GridTableLight", StringComparison.OrdinalIgnoreCase)) {
            canonicalStyleName = "TableGridLight";
            return true;
        }

        if (string.Equals(normalized, "PlainTable1", StringComparison.OrdinalIgnoreCase)) {
            canonicalStyleName = "PlainTable1";
            return true;
        }

        if (string.Equals(normalized, "GridTable1Light", StringComparison.OrdinalIgnoreCase)) {
            canonicalStyleName = "GridTable1Light";
            return true;
        }

        if (TryGetAccentNumber(normalized, "GridTable1LightAccent", out int gridAccent)) {
            canonicalStyleName = "GridTable1LightAccent" + gridAccent.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return true;
        }

        if (string.Equals(normalized, "ListTable1Light", StringComparison.OrdinalIgnoreCase)) {
            canonicalStyleName = "ListTable1Light";
            return true;
        }

        if (TryGetAccentNumber(normalized, "ListTable1LightAccent", out int listAccent)) {
            canonicalStyleName = "ListTable1LightAccent" + listAccent.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return true;
        }

        canonicalStyleName = null;
        return false;
    }

    private static ArgumentException CreateUnsupportedWordTableStyleException(string styleName) =>
        new ArgumentException(
            $"Unsupported Word table style '{styleName}'. Supported styles: {string.Join(", ", SupportedWordStyleNames)}.",
            nameof(styleName));

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

    private static bool TryGetAccentNumber(string normalized, string prefix, out int accentNumber) {
        accentNumber = 0;
        if (!normalized.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        string suffix = normalized.Substring(prefix.Length);
        if (suffix.Length != 1 || suffix[0] < '1' || suffix[0] > '6') {
            return false;
        }

        accentNumber = suffix[0] - '0';
        return true;
    }
}
