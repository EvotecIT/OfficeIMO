#pragma warning disable RS0027 // SetUnderline/SetStrikethrough(bool = true) mirror Word and PowerPoint beside their pattern overloads.
namespace OfficeIMO.Pdf;

public sealed partial class PdfTextRun {
    /// <summary>Makes this run bold or regular weight.</summary>
    public PdfTextRun SetBold(bool bold = true) { Bold = bold; return this; }

    /// <summary>Makes this run italic or upright.</summary>
    public PdfTextRun SetItalic(bool italic = true) { Italic = italic; return this; }

    /// <summary>Adds a single underline to this run, or removes its underline.</summary>
    public PdfTextRun SetUnderline(bool underline = true) => SetUnderline(underline ? OfficeIMO.Drawing.OfficeTextDecorationStyle.Single : OfficeIMO.Drawing.OfficeTextDecorationStyle.None);

    /// <summary>Sets the underline pattern of this run. <see cref="OfficeIMO.Drawing.OfficeTextDecorationStyle.None"/> removes it.</summary>
    public PdfTextRun SetUnderline(OfficeIMO.Drawing.OfficeTextDecorationStyle style) { ValidateUnderlineStyle(style, nameof(style)); UnderlineStyle = style; return this; }

    /// <summary>Adds a single strikethrough to this run, or removes its strikethrough.</summary>
    public PdfTextRun SetStrikethrough(bool strikethrough = true) => SetStrikethrough(strikethrough ? OfficeIMO.Drawing.OfficeTextDecorationStyle.Single : OfficeIMO.Drawing.OfficeTextDecorationStyle.None);

    /// <summary>Sets the strikethrough pattern of this run. <see cref="OfficeIMO.Drawing.OfficeTextDecorationStyle.None"/> removes it.</summary>
    public PdfTextRun SetStrikethrough(OfficeIMO.Drawing.OfficeTextDecorationStyle style) { ValidateStrikeStyle(style, nameof(style)); StrikeStyle = style; return this; }

    /// <summary>Sets the foreground colour of this run.</summary>
    public PdfTextRun SetColor(global::OfficeIMO.OfficeColor color) { Color = color; return this; }

    /// <summary>Sets the background (highlight) colour of this run.</summary>
    public PdfTextRun SetBackgroundColor(global::OfficeIMO.OfficeColor color) { BackgroundColor = color; return this; }

    /// <summary>Sets the font size of this run.</summary>
    public PdfTextRun SetFontSize(global::OfficeIMO.OfficeLength fontSize) { Guard.Positive(fontSize.ToPoints(), nameof(fontSize)); FontSize = fontSize; return this; }

    /// <summary>Sets the registered embedded font family of this run. <see cref="StandardFont"/> remains its fallback.</summary>
    public PdfTextRun SetFontFamily(string fontFamily) { Guard.NotNullOrWhiteSpace(fontFamily, nameof(fontFamily)); FontFamily = fontFamily.Trim(); return this; }

    /// <summary>Sets the standard PDF font of this run.</summary>
    public PdfTextRun SetStandardFont(PdfStandardFont standardFont) {
        Guard.StandardFont(standardFont, nameof(standardFont), "Text run font must be one of the supported standard PDF fonts.");
        StandardFont = standardFont;
        return this;
    }

    /// <summary>Sets the baseline placement of this run.</summary>
    public PdfTextRun SetBaseline(PdfTextBaseline baseline) { Guard.TextBaseline(baseline, nameof(baseline)); Baseline = baseline; return this; }

    /// <summary>Raises this run to superscript.</summary>
    public PdfTextRun SetSuperscript() => SetBaseline(PdfTextBaseline.Superscript);

    /// <summary>Lowers this run to subscript.</summary>
    public PdfTextRun SetSubscript() => SetBaseline(PdfTextBaseline.Subscript);

    /// <summary>Sets the horizontal glyph width percentage of this text run. 100 is normal width.</summary>
    public PdfTextRun SetHorizontalTextScaling(double percentage) {
        Guard.Positive(percentage, nameof(percentage));
        if (InlineElement == null) HorizontalTextScaling = percentage;
        return this;
    }

    /// <summary>Sets the additional advance after each glyph of this text run. Zero is normal spacing.</summary>
    public PdfTextRun SetCharacterSpacing(global::OfficeIMO.OfficeLength spacing) {
        if (InlineElement == null) CharacterSpacing = spacing;
        return this;
    }

    // Paragraph-created runs take the remaining defaults in place, so the run returned to the caller is the stored run.
    internal PdfTextRun ApplyParagraphDefaults(PdfTextRunDefaults defaults) {
        if (InlineElement != null) return this;
        FeatureSettings = defaults.FeatureSettings;
        HorizontalTextScaling = defaults.HorizontalTextScaling;
        CharacterSpacing = defaults.CharacterSpacing;
        return this;
    }

    internal static void ValidateUnderlineStyle(OfficeIMO.Drawing.OfficeTextDecorationStyle style, string parameterName) {
        if (style < OfficeIMO.Drawing.OfficeTextDecorationStyle.None || style > OfficeIMO.Drawing.OfficeTextDecorationStyle.Words) {
            throw new System.ArgumentOutOfRangeException(parameterName);
        }
    }

    internal static void ValidateStrikeStyle(OfficeIMO.Drawing.OfficeTextDecorationStyle style, string parameterName) {
        if (style < OfficeIMO.Drawing.OfficeTextDecorationStyle.None || style > OfficeIMO.Drawing.OfficeTextDecorationStyle.Wavy) {
            throw new System.ArgumentOutOfRangeException(parameterName);
        }
    }
}
